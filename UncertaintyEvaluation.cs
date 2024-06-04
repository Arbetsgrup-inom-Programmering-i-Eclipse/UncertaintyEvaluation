using System;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Windows.Forms;


[assembly: AssemblyVersion("3.0.0.10")]

[assembly: ESAPIScript(IsWriteable = true)]


namespace VMS.TPS
{
    public class Script
    {
        public Script()
        {

        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        public void Execute(ScriptContext context /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {
            PlanSetup planSetup = context.PlanSetup;
            // TODO : Add here the code that is called when the script is launched from Eclipse.
            if (planSetup == null)
                MessageBox.Show("Ingen plan vald!", "Check Field Size", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else if (planSetup.GetType() != typeof(IonPlanSetup))
                MessageBox.Show("Planen �r ej en protonplan!", "Check Field Size", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
                run(context.IonPlanSetup, context.Patient, context.Course);
        }

        public void run(IonPlanSetup plan, Patient patient, Course course)
        {

            patient.BeginModifications(); // F�r att kunna b�rja modifiera datan i patienten

            string MinplanID = plan.Id.Substring(0, 2) + "VoxlWiseMin"; // nya min plan ID
            string MaxplanID = plan.Id.Substring(0, 2) + "VoxlWiseMax"; // nya max plan ID
            bool ok = true;
            foreach (IonPlanSetup temp_plan in course.IonPlanSetups) // kollar om kursen redan inneh�ller en plan med det namnet.
            {
                if (temp_plan.Id == MinplanID || temp_plan.Id == MaxplanID)
                {
                    MessageBox.Show("Namnet UnEvalPlan �r redan upptaget, radera planen eller d�p om");
                    ok = false;
                }
            }
            if (ok)
            {
                ExternalPlanSetup planMin = course.AddExternalPlanSetup(plan.StructureSet); // Skapar en ny tom plan i samma kurs som 
                ExternalPlanSetup planMax = course.AddExternalPlanSetup(plan.StructureSet); // Skapar en ny tom plan i samma kurs som 
                planMin.Id = MinplanID; // D�per nya planen till VoxelWiseMin
                planMax.Id = MaxplanID; // D�per nya planen till VoxelWiseMax

                EvaluationDose evalDoseMin = planMin.CreateEvaluationDose(); // Skapar en tom dosmatris i den nya planen
                EvaluationDose evalDoseMax = planMax.CreateEvaluationDose(); // Skapar en tom dosmatris i den nya planen
                evalDoseMin = planMin.CopyEvaluationDose(plan.Dose); // kopierar dosen fr�n nominella planen f�r att f� samma dimensioner
                evalDoseMax = planMax.CopyEvaluationDose(plan.Dose); // kopierar dosen fr�n nominella planen f�r att f� samma dimensioner

                

                Dose dose = plan.Dose; // dosen
                int W, H, D; // Width, height, depth
                VVector rowDirection, columnDirection; //vilket h�ll pixlarna i matrisen �r lagda

                W = dose.XSize;
                H = dose.YSize;
                D = dose.ZSize;
                rowDirection = dose.XDirection; // x = rad
                columnDirection = dose.YDirection; // y = column

                List<int[,]> voxels = new List<int[,]>(); // lista �ver tomma dosmatriser f�r alla planer f�r ett plane
                bool doseMatrixSize = true;
                if (plan.PlanUncertainties.Count() > 0)
                {
                    List<PlanUncertainty> uPlans = new List<PlanUncertainty>(); // Eclipse f�r ibland f�r sig att skapa "sp�k-robustutv�rdering" som har ett space i namnet. Det h�r tar bort dem.
                    foreach (var uPlan in plan.PlanUncertainties)
                        if (!uPlan.Id.Contains(" "))
                        {
                            uPlans.Add(uPlan);
                            if (uPlan.Dose.XSize != W || uPlan.Dose.YSize != H || uPlan.Dose.ZSize != D || uPlan.Dose.Origin.x != plan.Dose.Origin.x || uPlan.Dose.Origin.y != plan.Dose.Origin.y || uPlan.Dose.Origin.z != plan.Dose.Origin.z) // Kontrollerar att dosmatrisen f�r nominella planen och alla uplaner �r samma storlek
                                doseMatrixSize = false;
                        }
                    if (doseMatrixSize)
                    {
                        voxels.Add(new int[dose.XSize, dose.YSize]); // skapar en tom dosmatris f�r nominella planens plane med voxlar med storleken fr�n dosmatrisen
                        foreach (var k in uPlans) // lopar genom alla robustutv�rderingar
                            voxels.Add(new int[dose.XSize, dose.YSize]); // robustutv�rderingsplanens plane med voxlar med storleken fr�n dosmatrisen

                        for (int z = 0; z < D; z++) // lopar genom alla planes
                        {
                            int[,] evalVoxelsMin = new int[evalDoseMin.XSize, evalDoseMin.YSize]; // skapar en tom dosmatris f�r VoxelWiseMin
                            int[,] evalVoxelsMax = new int[evalDoseMax.XSize, evalDoseMax.YSize]; // skapar en tom dosmatris f�r VoxelWiseMax
                            dose.GetVoxels(z, voxels.First()); // h�mtar alla voxelv�rden fr�n nominella planen
                            for (int k = 0; k < voxels.Count - 1; k++) //voxels.count - 1 f�r att undvika nominella planen
                                uPlans.ElementAt(k).Dose.GetVoxels(z, voxels.ElementAt(k + 1)); // h�mtar alla voxelv�rden fr�n robustutv�rderingsplanerna en i taget, k �r k:te uplan medan voxels �r k+1 eftersom f�rsta �r nominella

                            for (int y = 0; y < evalDoseMin.YSize; y++) // loopar genom alla y 
                            {
                                for (int x = 0; x < evalDoseMin.XSize; x++) // loopar genom alla x
                                {
                                    int voxelValueMin = voxels.First()[x, y]; // tar ut voxelv�rdet f�r x och y fr�n nominella planen
                                    int voxelValueMax = voxels.First()[x, y];
                                    for (int l = 1; l < voxels.Count; l++) // loopar genom alla uPlan-dosmatriser och j�mf�r mot v�rdet fr�n nominella
                                    {
                                        if (voxelValueMin > voxels.ElementAt(l)[x, y]) // om v�rdet �r mindre �n det tidigare, ers�tt
                                            voxelValueMin = voxels.ElementAt(l)[x, y];
                                        if (voxelValueMax < voxels.ElementAt(l)[x, y]) // om v�rdet �r h�gre �n det tidigare, ers�tt
                                            voxelValueMax = voxels.ElementAt(l)[x, y];
                                    }
                                    evalVoxelsMin[x, y] = voxelValueMin; // ans�tter det minsta v�rdet till VoxelWiseMin voxel.
                                    evalVoxelsMax[x, y] = voxelValueMax; // ans�tter det h�gsta v�rdet till VoxelWiseMax voxel.
                                }
                            }
                            evalDoseMin.SetVoxels(z, evalVoxelsMin); // Ans�tter nya voxelv�rdena till dosmatrisen f�r VoxelWiseMin
                            evalDoseMax.SetVoxels(z, evalVoxelsMax); // Ans�tter nya voxelv�rdena till dosmatrisen f�r VoxelWiseMax
                        }
                        MessageBox.Show("Gl�m inte byta normering till: " + plan.PlanNormalizationValue.ToString()); // Skriver ut normeringsv�rdet fr�n nominella planen f�r att p�minna anv�ndaren att normera om de skapade planerna innan granskning.
                        MessageBox.Show("Done!");
                    }
                    else
                        MessageBox.Show("Dosmatrisen var inte samma storlek i nominella planen som i n�gon robustutv�rdering");
                }
                else
                    MessageBox.Show("Finns inga os�kerhetsber�kningar");
            }
        }
    }
}
