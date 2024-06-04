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
                MessageBox.Show("Planen är ej en protonplan!", "Check Field Size", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
                run(context.IonPlanSetup, context.Patient, context.Course);
        }

        public void run(IonPlanSetup plan, Patient patient, Course course)
        {

            patient.BeginModifications(); // För att kunna börja modifiera datan i patienten

            string MinplanID = plan.Id.Substring(0, 2) + "VoxlWiseMin"; // nya min plan ID
            string MaxplanID = plan.Id.Substring(0, 2) + "VoxlWiseMax"; // nya max plan ID
            bool ok = true;
            foreach (IonPlanSetup temp_plan in course.IonPlanSetups) // kollar om kursen redan innehåller en plan med det namnet.
            {
                if (temp_plan.Id == MinplanID || temp_plan.Id == MaxplanID)
                {
                    MessageBox.Show("Namnet UnEvalPlan är redan upptaget, radera planen eller döp om");
                    ok = false;
                }
            }
            if (ok)
            {
                ExternalPlanSetup planMin = course.AddExternalPlanSetup(plan.StructureSet); // Skapar en ny tom plan i samma kurs som 
                ExternalPlanSetup planMax = course.AddExternalPlanSetup(plan.StructureSet); // Skapar en ny tom plan i samma kurs som 
                planMin.Id = MinplanID; // Döper nya planen till VoxelWiseMin
                planMax.Id = MaxplanID; // Döper nya planen till VoxelWiseMax

                EvaluationDose evalDoseMin = planMin.CreateEvaluationDose(); // Skapar en tom dosmatris i den nya planen
                EvaluationDose evalDoseMax = planMax.CreateEvaluationDose(); // Skapar en tom dosmatris i den nya planen
                evalDoseMin = planMin.CopyEvaluationDose(plan.Dose); // kopierar dosen från nominella planen för att få samma dimensioner
                evalDoseMax = planMax.CopyEvaluationDose(plan.Dose); // kopierar dosen från nominella planen för att få samma dimensioner

                

                Dose dose = plan.Dose; // dosen
                int W, H, D; // Width, height, depth
                VVector rowDirection, columnDirection; //vilket håll pixlarna i matrisen är lagda

                W = dose.XSize;
                H = dose.YSize;
                D = dose.ZSize;
                rowDirection = dose.XDirection; // x = rad
                columnDirection = dose.YDirection; // y = column

                List<int[,]> voxels = new List<int[,]>(); // lista över tomma dosmatriser för alla planer för ett plane
                bool doseMatrixSize = true;
                if (plan.PlanUncertainties.Count() > 0)
                {
                    List<PlanUncertainty> uPlans = new List<PlanUncertainty>(); // Eclipse får ibland för sig att skapa "spök-robustutvärdering" som har ett space i namnet. Det här tar bort dem.
                    foreach (var uPlan in plan.PlanUncertainties)
                        if (!uPlan.Id.Contains(" "))
                        {
                            uPlans.Add(uPlan);
                            if (uPlan.Dose.XSize != W || uPlan.Dose.YSize != H || uPlan.Dose.ZSize != D || uPlan.Dose.Origin.x != plan.Dose.Origin.x || uPlan.Dose.Origin.y != plan.Dose.Origin.y || uPlan.Dose.Origin.z != plan.Dose.Origin.z) // Kontrollerar att dosmatrisen för nominella planen och alla uplaner är samma storlek
                                doseMatrixSize = false;
                        }
                    if (doseMatrixSize)
                    {
                        voxels.Add(new int[dose.XSize, dose.YSize]); // skapar en tom dosmatris för nominella planens plane med voxlar med storleken från dosmatrisen
                        foreach (var k in uPlans) // lopar genom alla robustutvärderingar
                            voxels.Add(new int[dose.XSize, dose.YSize]); // robustutvärderingsplanens plane med voxlar med storleken från dosmatrisen

                        for (int z = 0; z < D; z++) // lopar genom alla planes
                        {
                            int[,] evalVoxelsMin = new int[evalDoseMin.XSize, evalDoseMin.YSize]; // skapar en tom dosmatris för VoxelWiseMin
                            int[,] evalVoxelsMax = new int[evalDoseMax.XSize, evalDoseMax.YSize]; // skapar en tom dosmatris för VoxelWiseMax
                            dose.GetVoxels(z, voxels.First()); // hämtar alla voxelvärden från nominella planen
                            for (int k = 0; k < voxels.Count - 1; k++) //voxels.count - 1 för att undvika nominella planen
                                uPlans.ElementAt(k).Dose.GetVoxels(z, voxels.ElementAt(k + 1)); // hämtar alla voxelvärden från robustutvärderingsplanerna en i taget, k är k:te uplan medan voxels är k+1 eftersom första är nominella

                            for (int y = 0; y < evalDoseMin.YSize; y++) // loopar genom alla y 
                            {
                                for (int x = 0; x < evalDoseMin.XSize; x++) // loopar genom alla x
                                {
                                    int voxelValueMin = voxels.First()[x, y]; // tar ut voxelvärdet för x och y från nominella planen
                                    int voxelValueMax = voxels.First()[x, y];
                                    for (int l = 1; l < voxels.Count; l++) // loopar genom alla uPlan-dosmatriser och jämför mot värdet från nominella
                                    {
                                        if (voxelValueMin > voxels.ElementAt(l)[x, y]) // om värdet är mindre än det tidigare, ersätt
                                            voxelValueMin = voxels.ElementAt(l)[x, y];
                                        if (voxelValueMax < voxels.ElementAt(l)[x, y]) // om värdet är högre än det tidigare, ersätt
                                            voxelValueMax = voxels.ElementAt(l)[x, y];
                                    }
                                    evalVoxelsMin[x, y] = voxelValueMin; // ansätter det minsta värdet till VoxelWiseMin voxel.
                                    evalVoxelsMax[x, y] = voxelValueMax; // ansätter det högsta värdet till VoxelWiseMax voxel.
                                }
                            }
                            evalDoseMin.SetVoxels(z, evalVoxelsMin); // Ansätter nya voxelvärdena till dosmatrisen för VoxelWiseMin
                            evalDoseMax.SetVoxels(z, evalVoxelsMax); // Ansätter nya voxelvärdena till dosmatrisen för VoxelWiseMax
                        }
                        MessageBox.Show("Glöm inte byta normering till: " + plan.PlanNormalizationValue.ToString()); // Skriver ut normeringsvärdet från nominella planen för att påminna användaren att normera om de skapade planerna innan granskning.
                        MessageBox.Show("Done!");
                    }
                    else
                        MessageBox.Show("Dosmatrisen var inte samma storlek i nominella planen som i någon robustutvärdering");
                }
                else
                    MessageBox.Show("Finns inga osäkerhetsberäkningar");
            }
        }
    }
}
