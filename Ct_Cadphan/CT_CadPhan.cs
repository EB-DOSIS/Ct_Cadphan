///////////////////////////////////////////////////////////////////////////////////////////////////
// Created 21.10.2020 Eyck Blank
// nachdem Fehler in HuCheckCadPhan gefunden
// und sich OxyPlot und WPF nicht mehr herauslösen liessen
// 12.11.2020 Einfügen der z_rand_X
//
///////////////////////////////////////////////////////////////////////////////////////////////////


using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.CompilerServices;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Drawing;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;
using System.Resources;


// TODO: Replace the following version attributes by creating AssemblyInfo.cs. You can do this in the properties of the Visual Studio project.
[assembly: AssemblyVersion("1.0.0.3")]
[assembly: AssemblyFileVersion("1.0.0.3")]
[assembly: AssemblyInformationalVersion("1.03")]

// TODO: Uncomment the following line if the script requires write access.
 [assembly: ESAPIScript(IsWriteable = true)]

namespace VMS.TPS
{
    public class Script
    {
        public Script()
        {
        }


        // Change these IDs to match your clinical conventions 
        const string BODY_ID = "BODY";
        const string AIR_ID = "Air";
        const string ACRYLIC_ID = "Acrylic";
        const string DELDRINE_ID = "Deldrine";
        const string LDPE_ID = "LDPE";
        const string PMP_ID = "PMP";
        const string POLYSTYRENE_ID = "PolyStyrene";
        const string TEFLON_ID = "Teflon";
        const string SCRIPT_NAME = "Opt Structures Script";

        const string TRANS_LO_ID = "Trans_LO";
        const string TRANS_RO_ID = "Trans_RO";
        const string TRANS_LU_ID = "Trans_LU";
        const string TRANS_RU_ID = "Trans_RU";

        const string Z1_L_ID = "Z1_L";
        const string Z1_O_ID = "Z1_O";
        const string Z1_R_ID = "Z1_R";
        const string Z1_U_ID = "Z1_U";

        const string Z2_L_ID = "Z2_L";
        const string Z2_O_ID = "Z2_O";
        const string Z2_R_ID = "Z2_R";
        const string Z2_U_ID = "Z2_U";

        const string Z3_L_ID = "Z3_L";
        const string Z3_O_ID = "Z3_O";
        const string Z3_R_ID = "Z3_R";
        const string Z3_U_ID = "Z3_U";

        const string Z4_L_ID = "Z4_L";
        const string Z4_O_ID = "Z4_O";
        const string Z4_R_ID = "Z4_R";
        const string Z4_U_ID = "Z4_U";

        const string Z_RAND_1_ID = "Z_Rand_1";
        const string Z_RAND_2_ID = "Z_Rand_2";
        const string Z_RAND_3_ID = "Z_Rand_3";
        const string Z_RAND_4_ID = "Z_Rand_4";

        const string MTF_0_ID = "MTF_0";


        const string UNIF_CENT_ID = "unif_cent";

        const string UNIF_AO_ID = "unif_ao";
        const string UNIF_IO_ID = "unif_io";

        const string UNIF_AR_ID = "unif_ar";
        const string UNIF_IR_ID = "unif_ir";

        const string UNIF_AU_ID = "unif_au";
        const string UNIF_IU_ID = "unif_iu";

        const string UNIF_AL_ID = "unif_al";
        const string UNIF_IL_ID = "unif_il";


        //==================================================
        // Berechnung der HU-Mittelwerte eines 3x3x3 Würfels
        //==================================================
        public void Execute(ScriptContext context, System.Windows.Window window /*, System.Windows.Window window, ScriptEnvironment environment*/)
        {
            double MittelwertAuslesen(int sIx, int sIy, int sIz, VMS.TPS.Common.Model.API.Image bild)
            {
                int[,] bildPlane = new int[bild.XSize, bild.YSize];
                bild.GetVoxels(sIz - 1, bildPlane);
                double sIV1 = bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy - 1]) + bild.VoxelToDisplayValue(bildPlane[sIx, sIy - 1]) + bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy - 1])
                    + bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy]) + bild.VoxelToDisplayValue(bildPlane[sIx, sIy]) + bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy])
                    + bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy + 1]) + bild.VoxelToDisplayValue(bildPlane[sIx, sIy + 1]) + bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy + 1]);
                bild.GetVoxels(sIz, bildPlane);
                double sIV2 = bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy - 1]) + bild.VoxelToDisplayValue(bildPlane[sIx, sIy - 1]) + bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy - 1])
                    + bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy]) + bild.VoxelToDisplayValue(bildPlane[sIx, sIy]) + bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy])
                    + bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy + 1]) + bild.VoxelToDisplayValue(bildPlane[sIx, sIy + 1]) + bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy + 1]);
                bild.GetVoxels(sIz + 1, bildPlane);
                double sIV3 = bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy - 1]) + bild.VoxelToDisplayValue(bildPlane[sIx, sIy - 1]) + bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy - 1])
                    + bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy]) + bild.VoxelToDisplayValue(bildPlane[sIx, sIy]) + bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy])
                    + bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy + 1]) + bild.VoxelToDisplayValue(bildPlane[sIx, sIy + 1]) + bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy + 1]);
                double sIV = (sIV1 + sIV2 + sIV3) / 27;
                return sIV;
            }

            double StreuungAuslesen(int sIx, int sIy, int sIz, VMS.TPS.Common.Model.API.Image bild)
            {
                double[] streuung = new double[27];
                int[,] bildPlane = new int[bild.XSize, bild.YSize];
                bild.GetVoxels(sIz - 1, bildPlane);
                streuung[0] = bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy - 1]);
                streuung[1] = bild.VoxelToDisplayValue(bildPlane[sIx, sIy - 1]);
                streuung[2] = bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy - 1]);
                streuung[3] = bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy]);
                streuung[4] = bild.VoxelToDisplayValue(bildPlane[sIx, sIy]);
                streuung[5] = bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy]);
                streuung[6] = bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy + 1]);
                streuung[7] = bild.VoxelToDisplayValue(bildPlane[sIx, sIy + 1]);
                streuung[8] = bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy + 1]);
                bild.GetVoxels(sIz, bildPlane);
                streuung[9] = bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy - 1]);
                streuung[10] = bild.VoxelToDisplayValue(bildPlane[sIx, sIy - 1]);
                streuung[11] = bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy - 1]);
                streuung[12] = bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy]);
                streuung[13] = bild.VoxelToDisplayValue(bildPlane[sIx, sIy]);
                streuung[14] = bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy]);
                streuung[15] = bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy + 1]);
                streuung[16] = bild.VoxelToDisplayValue(bildPlane[sIx, sIy + 1]);
                streuung[17] = bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy + 1]);
                bild.GetVoxels(sIz + 1, bildPlane);
                streuung[18] = bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy - 1]);
                streuung[19] = bild.VoxelToDisplayValue(bildPlane[sIx, sIy - 1]);
                streuung[20] = bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy - 1]);
                streuung[21] = bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy]);
                streuung[22] = bild.VoxelToDisplayValue(bildPlane[sIx, sIy]);
                streuung[23] = bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy]);
                streuung[24] = bild.VoxelToDisplayValue(bildPlane[sIx - 1, sIy + 1]);
                streuung[25] = bild.VoxelToDisplayValue(bildPlane[sIx, sIy + 1]);
                streuung[26] = bild.VoxelToDisplayValue(bildPlane[sIx + 1, sIy + 1]);

                double average = streuung.Average();
                double sumOfSquaresOfDifferences = streuung.Select(val => (val - average) * (val - average)).Sum();
                double sSV = Math.Sqrt(sumOfSquaresOfDifferences / streuung.Length);

                return sSV;
            }



            //++++++++++++++++++++++++++++++++++++++++++++++++++


            if (context.Patient == null || context.StructureSet == null)
            {
                MessageBox.Show("Please load a patient, 3D image, and structure set before running this script.", SCRIPT_NAME, MessageBoxButton.OKCancel, MessageBoxImage.Exclamation);
                return;
            }
            StructureSet ss = context.StructureSet;

            //=========================
            // Auffinden der Strukturen
            //=========================
            // find Body 
            Structure body = ss.Structures.FirstOrDefault(x => x.Id == BODY_ID);
            if (body == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", BODY_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Air 
            Structure air = ss.Structures.FirstOrDefault(x => x.Id == AIR_ID);
            if (air == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", AIR_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Acrylic 
            Structure acrylic = ss.Structures.FirstOrDefault(x => x.Id == ACRYLIC_ID);
            if (acrylic == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", ACRYLIC_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Deldrine 
            Structure deldrine = ss.Structures.FirstOrDefault(x => x.Id == DELDRINE_ID);
            if (deldrine == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", DELDRINE_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Ldpe 
            Structure ldpe = ss.Structures.FirstOrDefault(x => x.Id == LDPE_ID);
            if (ldpe == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", LDPE_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Pmp 
            Structure pmp = ss.Structures.FirstOrDefault(x => x.Id == PMP_ID);
            if (pmp == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", PMP_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find PolyStyrene 
            Structure polystyrene = ss.Structures.FirstOrDefault(x => x.Id == POLYSTYRENE_ID);
            if (polystyrene == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", POLYSTYRENE_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Teflon 
            Structure teflon = ss.Structures.FirstOrDefault(x => x.Id == TEFLON_ID);
            if (teflon == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", TEFLON_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }



            // find Trans_LO 
            Structure trans_lo = ss.Structures.FirstOrDefault(x => x.Id == TRANS_LO_ID);
            if (trans_lo == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", TRANS_LO_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Trans_RO 
            Structure trans_ro = ss.Structures.FirstOrDefault(x => x.Id == TRANS_RO_ID);
            if (trans_ro == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", TRANS_RO_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Trans_LU 
            Structure trans_lu = ss.Structures.FirstOrDefault(x => x.Id == TRANS_LU_ID);
            if (trans_lu == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", TRANS_LU_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Trans_RU 
            Structure trans_ru = ss.Structures.FirstOrDefault(x => x.Id == TRANS_RU_ID);
            if (trans_ru == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", TRANS_RU_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }


            // find Z1_L 
            Structure z1_l = ss.Structures.FirstOrDefault(x => x.Id == Z1_L_ID);
            if (z1_l == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z1_L_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z1_O 
            Structure z1_o = ss.Structures.FirstOrDefault(x => x.Id == Z1_O_ID);
            if (z1_o == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z1_O_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z1_R 
            Structure z1_r = ss.Structures.FirstOrDefault(x => x.Id == Z1_R_ID);
            if (z1_r == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z1_R_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z1_U 
            Structure z1_u = ss.Structures.FirstOrDefault(x => x.Id == Z1_U_ID);
            if (z1_u == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z1_U_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z2_L 
            Structure z2_l = ss.Structures.FirstOrDefault(x => x.Id == Z2_L_ID);
            if (z2_l == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z2_L_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z2_O 
            Structure z2_o = ss.Structures.FirstOrDefault(x => x.Id == Z2_O_ID);
            if (z2_o == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z2_O_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z2_R 
            Structure z2_r = ss.Structures.FirstOrDefault(x => x.Id == Z2_R_ID);
            if (z2_r == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z2_R_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z2_U 
            Structure z2_u = ss.Structures.FirstOrDefault(x => x.Id == Z2_U_ID);
            if (z2_u == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z2_U_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z3_L 
            Structure z3_l = ss.Structures.FirstOrDefault(x => x.Id == Z3_L_ID);
            if (z3_l == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z3_L_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z3_O 
            Structure z3_o = ss.Structures.FirstOrDefault(x => x.Id == Z3_O_ID);
            if (z3_o == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z3_O_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z3_R 
            Structure z3_r = ss.Structures.FirstOrDefault(x => x.Id == Z3_R_ID);
            if (z3_r == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z3_R_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z3_U 
            Structure z3_u = ss.Structures.FirstOrDefault(x => x.Id == Z3_U_ID);
            if (z3_u == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z3_U_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z4_L 
            Structure z4_l = ss.Structures.FirstOrDefault(x => x.Id == Z4_L_ID);
            if (z4_l == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z4_L_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z4_O 
            Structure z4_o = ss.Structures.FirstOrDefault(x => x.Id == Z4_O_ID);
            if (z4_o == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z4_O_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z4_R 
            Structure z4_r = ss.Structures.FirstOrDefault(x => x.Id == Z4_R_ID);
            if (z4_r == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z4_R_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z4_U 
            Structure z4_u = ss.Structures.FirstOrDefault(x => x.Id == Z4_U_ID);
            if (z4_u == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z4_U_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find MTF_0 
            Structure mtf_0 = ss.Structures.FirstOrDefault(x => x.Id == MTF_0_ID);
            if (mtf_0 == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", MTF_0_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }


            // find Z_rand_1 
            Structure z_rand_1 = ss.Structures.FirstOrDefault(x => x.Id == Z_RAND_1_ID);
            if (z_rand_1 == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z_RAND_1_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z_rand_2 
            Structure z_rand_2 = ss.Structures.FirstOrDefault(x => x.Id == Z_RAND_2_ID);
            if (z_rand_2 == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z_RAND_2_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z_rand_3 
            Structure z_rand_3 = ss.Structures.FirstOrDefault(x => x.Id == Z_RAND_3_ID);
            if (z_rand_3 == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z_RAND_3_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find Z_rand_4 
            Structure z_rand_4 = ss.Structures.FirstOrDefault(x => x.Id == Z_RAND_4_ID);
            if (z_rand_4 == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", Z_RAND_4_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }


            // Uniformitätstest
            // find unif_cent
            Structure unif_cent = ss.Structures.FirstOrDefault(x => x.Id == UNIF_CENT_ID);
            if (unif_cent == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", UNIF_CENT_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find unif_ao
            Structure unif_ao = ss.Structures.FirstOrDefault(x => x.Id == UNIF_AO_ID);
            if (unif_ao == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", UNIF_AO_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            // find unif_io
            Structure unif_io = ss.Structures.FirstOrDefault(x => x.Id == UNIF_IO_ID);
            if (unif_io == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", UNIF_IO_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find unif_ar
            Structure unif_ar = ss.Structures.FirstOrDefault(x => x.Id == UNIF_AR_ID);
            if (unif_ar == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", UNIF_AR_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            // find unif_ir
            Structure unif_ir = ss.Structures.FirstOrDefault(x => x.Id == UNIF_IR_ID);
            if (unif_ir == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", UNIF_IR_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find unif_au
            Structure unif_au = ss.Structures.FirstOrDefault(x => x.Id == UNIF_AU_ID);
            if (unif_au == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", UNIF_AU_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            // find unif_iu
            Structure unif_iu = ss.Structures.FirstOrDefault(x => x.Id == UNIF_IU_ID);
            if (unif_iu == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", UNIF_IU_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            // find unif_al
            Structure unif_al = ss.Structures.FirstOrDefault(x => x.Id == UNIF_AL_ID);
            if (unif_al == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", UNIF_AL_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }
            // find unif_il
            Structure unif_il = ss.Structures.FirstOrDefault(x => x.Id == UNIF_IL_ID);
            if (unif_il == null)
            {
                MessageBox.Show(string.Format("'{0}' not found!", UNIF_IL_ID), SCRIPT_NAME, MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }


            //=======================================
            // Berechnen der Transformationsparameter
            //=======================================
            context.Patient.BeginModifications();   // enable writing with this script. 

            VMS.TPS.Common.Model.API.Image image = context.Image;
            String Datum = image.CreationDateTime.Value.ToString("dd.MM.yyyy");
            String Zeit = image.CreationDateTime.Value.ToString("HH:mm:ss");

            int SizeIX = image.XSize;
            int SizeIY = image.YSize;
            int SizeIZ = image.YSize;

            double uoX = image.UserOrigin.x;
            double uoY = image.UserOrigin.y;
            double uoZ = image.UserOrigin.z;

            double oX = image.Origin.x;
            double oY = image.Origin.y;
            double oZ = image.Origin.z;

            double sizeX = image.XSize;
            double sizeY = image.YSize;
            double sizeZ = image.ZSize;

            double resX = image.XRes;
            double resY = image.YRes;
            double resZ = image.ZRes;


            if (image == null)
            {
                MessageBox.Show("Please load a 3D image.", "VarianDeveloper");
                return;
            }

            //============================ 
            // Center HU Body 
            //============================ 
            // Center
            double bodyCx = body.CenterPoint.x;
            double bodyCy = body.CenterPoint.y;
            double bodyCz = body.CenterPoint.z;
            // HU

            //============================ 
            // Center HU Air 
            //============================ 
            // Center
            double airCx = air.CenterPoint.x;
            double airCy = air.CenterPoint.y;
            double airCz = air.CenterPoint.z;
            // Umrechnung der cm in VoxelIndex
            int airIx = Convert.ToInt32(Math.Ceiling((air.CenterPoint.x - oX) / resX));
            int airIy = Convert.ToInt32(Math.Ceiling((air.CenterPoint.y - oY) / resY));
            int airIz = Convert.ToInt32(Math.Ceiling((air.CenterPoint.z - oZ) / resZ));
            // HU
            double airIV = MittelwertAuslesen(airIx, airIy, airIz, image);


            //============================ 
            // Center HU Acrylic 
            //============================ 
            // Center   
            double acrylicCx = acrylic.CenterPoint.x;
            double acrylicCy = acrylic.CenterPoint.y;
            double acrylicCz = acrylic.CenterPoint.z;
            // Umrechnung der cm in VoxelIndex
            int acrylicIx = Convert.ToInt32(Math.Ceiling((acrylic.CenterPoint.x - oX) / resX));
            int acrylicIy = Convert.ToInt32(Math.Ceiling((acrylic.CenterPoint.y - oY) / resY));
            int acrylicIz = Convert.ToInt32(Math.Ceiling((acrylic.CenterPoint.z - oZ) / resZ));
            // HU
            double acrylicIV = MittelwertAuslesen(acrylicIx, acrylicIy, acrylicIz, image);

            //============================ 
            // Center HU Deldrine 
            //============================ 
            // Center
            double deldrineCx = deldrine.CenterPoint.x;
            double deldrineCy = deldrine.CenterPoint.y;
            double deldrineCz = deldrine.CenterPoint.z;
            // Umrechnung der cm in VoxelIndex
            int deldrineIx = Convert.ToInt32(Math.Ceiling((deldrine.CenterPoint.x - oX) / resX));
            int deldrineIy = Convert.ToInt32(Math.Ceiling((deldrine.CenterPoint.y - oY) / resY));
            int deldrineIz = Convert.ToInt32(Math.Ceiling((deldrine.CenterPoint.z - oZ) / resZ));
            // HU
            double deldrineIV = MittelwertAuslesen(deldrineIx, deldrineIy, deldrineIz, image);

            //============================ 
            // Center HU LDPE 
            //============================ 
            // Center
            double ldpeCx = ldpe.CenterPoint.x;
            double ldpeCy = ldpe.CenterPoint.y;
            double ldpeCz = ldpe.CenterPoint.z;
            // Umrechnung der cm in VoxelIndex
            int ldpeIx = Convert.ToInt32(Math.Ceiling((ldpe.CenterPoint.x - oX) / resX));
            int ldpeIy = Convert.ToInt32(Math.Ceiling((ldpe.CenterPoint.y - oY) / resY));
            int ldpeIz = Convert.ToInt32(Math.Ceiling((ldpe.CenterPoint.z - oZ) / resZ));
            // HU
            double ldpeIV = MittelwertAuslesen(ldpeIx, ldpeIy, ldpeIz, image);

            //============================ 
            // Center HU PMP 
            //============================ 
            // Center
            double pmpCx = pmp.CenterPoint.x;
            double pmpCy = pmp.CenterPoint.y;
            double pmpCz = pmp.CenterPoint.z;
            // Umrechnung der cm in VoxelIndex
            int pmpIx = Convert.ToInt32(Math.Ceiling((pmp.CenterPoint.x - oX) / resX));
            int pmpIy = Convert.ToInt32(Math.Ceiling((pmp.CenterPoint.y - oY) / resY));
            int pmpIz = Convert.ToInt32(Math.Ceiling((pmp.CenterPoint.z - oZ) / resZ));
            // HU
            double pmpIV = MittelwertAuslesen(pmpIx, pmpIy, pmpIz, image);

            //============================ 
            // Center HU PolyStyrene 
            //============================ 
            // Center
            double polystyreneCx = polystyrene.CenterPoint.x;
            double polystyreneCy = polystyrene.CenterPoint.y;
            double polystyreneCz = polystyrene.CenterPoint.z;
            // Umrechnung der cm in VoxelIndex
            int polystyreneIx = Convert.ToInt32(Math.Ceiling((polystyrene.CenterPoint.x - oX) / resX));
            int polystyreneIy = Convert.ToInt32(Math.Ceiling((polystyrene.CenterPoint.y - oY) / resY));
            int polystyreneIz = Convert.ToInt32(Math.Ceiling((polystyrene.CenterPoint.z - oZ) / resZ));
            // HU
            double polystyreneIV = MittelwertAuslesen(polystyreneIx, polystyreneIy, polystyreneIz, image);

            //============================ 
            // Center HU Teflon 
            //============================ 
            // Center
            double teflonCx = teflon.CenterPoint.x;
            double teflonCy = teflon.CenterPoint.y;
            double teflonCz = teflon.CenterPoint.z;
            // Umrechnung der cm in VoxelIndex
            int teflonIx = Convert.ToInt32(Math.Ceiling((teflon.CenterPoint.x - oX) / resY));
            int teflonIy = Convert.ToInt32(Math.Ceiling((teflon.CenterPoint.y - oY) / resY));
            int teflonIz = Convert.ToInt32(Math.Ceiling((teflon.CenterPoint.z - oZ) / resZ));
            // HU
            double teflonIV = MittelwertAuslesen(teflonIx, teflonIy, teflonIz, image);


            //===================== 
            // Center Trans_LO 
            //===================== 
            double LoCx = trans_lo.CenterPoint.x;
            double LoCy = trans_lo.CenterPoint.y;
            double LoCz = trans_lo.CenterPoint.z;

            //===================== 
            // Center Trans_RO 
            //===================== 
            double RoCx = trans_ro.CenterPoint.x;
            double RoCy = trans_ro.CenterPoint.y;
            double RoCz = trans_ro.CenterPoint.z;

            //===================== 
            // Center Trans_LU 
            //===================== 
            double LuCx = trans_lu.CenterPoint.x;
            double LuCy = trans_lu.CenterPoint.y;
            double LuCz = trans_lu.CenterPoint.z;

            //===================== 
            // Center Trans_RU 
            //===================== 
            double RuCx = trans_ru.CenterPoint.x;
            double RuCy = trans_ru.CenterPoint.y;
            double RuCz = trans_ru.CenterPoint.z;


            //===================== 
            // Center z1 
            //===================== 
            double z1_L_Cx = z1_l.CenterPoint.x;
            double z1_L_Cy = z1_l.CenterPoint.y;
            double z1_L_Cz = z1_l.CenterPoint.z;

            double z1_O_Cx = z1_o.CenterPoint.x;
            double z1_O_Cy = z1_o.CenterPoint.y;
            double z1_O_Cz = z1_o.CenterPoint.z;

            double z1_R_Cx = z1_r.CenterPoint.x;
            double z1_R_Cy = z1_r.CenterPoint.y;
            double z1_R_Cz = z1_r.CenterPoint.z;

            double z1_U_Cx = z1_u.CenterPoint.x;
            double z1_U_Cy = z1_u.CenterPoint.y;
            double z1_U_Cz = z1_u.CenterPoint.z;

            //===================== 
            // Center z2 
            //===================== 
            double z2_L_Cx = z2_l.CenterPoint.x;
            double z2_L_Cy = z2_l.CenterPoint.y;
            double z2_L_Cz = z2_l.CenterPoint.z;

            double z2_O_Cx = z2_o.CenterPoint.x;
            double z2_O_Cy = z2_o.CenterPoint.y;
            double z2_O_Cz = z2_o.CenterPoint.z;

            double z2_R_Cx = z2_r.CenterPoint.x;
            double z2_R_Cy = z2_r.CenterPoint.y;
            double z2_R_Cz = z2_r.CenterPoint.z;

            double z2_U_Cx = z2_u.CenterPoint.x;
            double z2_U_Cy = z2_u.CenterPoint.y;
            double z2_U_Cz = z2_u.CenterPoint.z;

            //===================== 
            // Center z3 
            //===================== 
            double z3_L_Cx = z3_l.CenterPoint.x;
            double z3_L_Cy = z3_l.CenterPoint.y;
            double z3_L_Cz = z3_l.CenterPoint.z;

            double z3_O_Cx = z3_o.CenterPoint.x;
            double z3_O_Cy = z3_o.CenterPoint.y;
            double z3_O_Cz = z3_o.CenterPoint.z;

            double z3_R_Cx = z3_r.CenterPoint.x;
            double z3_R_Cy = z3_r.CenterPoint.y;
            double z3_R_Cz = z3_r.CenterPoint.z;

            double z3_U_Cx = z3_u.CenterPoint.x;
            double z3_U_Cy = z3_u.CenterPoint.y;
            double z3_U_Cz = z3_u.CenterPoint.z;

            //===================== 
            // Center z4 
            //===================== 
            double z4_L_Cx = z4_l.CenterPoint.x;
            double z4_L_Cy = z4_l.CenterPoint.y;
            double z4_L_Cz = z4_l.CenterPoint.z;

            double z4_O_Cx = z4_o.CenterPoint.x;
            double z4_O_Cy = z4_o.CenterPoint.y;
            double z4_O_Cz = z4_o.CenterPoint.z;

            double z4_R_Cx = z4_r.CenterPoint.x;
            double z4_R_Cy = z4_r.CenterPoint.y;
            double z4_R_Cz = z4_r.CenterPoint.z;

            double z4_U_Cx = z4_u.CenterPoint.x;
            double z4_U_Cy = z4_u.CenterPoint.y;
            double z4_U_Cz = z4_u.CenterPoint.z;

            //===================== 
            // Center z_rand_X 
            //===================== 

            double z_rand_1_Cx = z_rand_1.CenterPoint.x;
            double z_rand_1_Cy = z_rand_1.CenterPoint.y;
            double z_rand_1_Cz = z_rand_1.CenterPoint.z;

            double z_rand_2_Cx = z_rand_2.CenterPoint.x;
            double z_rand_2_Cy = z_rand_2.CenterPoint.y;
            double z_rand_2_Cz = z_rand_2.CenterPoint.z;

            double z_rand_3_Cx = z_rand_3.CenterPoint.x;
            double z_rand_3_Cy = z_rand_3.CenterPoint.y;
            double z_rand_3_Cz = z_rand_3.CenterPoint.z;

            double z_rand_4_Cx = z_rand_4.CenterPoint.x;
            double z_rand_4_Cy = z_rand_4.CenterPoint.y;
            double z_rand_4_Cz = z_rand_4.CenterPoint.z;


            //===================== 
            // Masse o u l r 
            //=====================

            double mass_o = RoCx - LoCx;
            double mass_u = RuCx - LuCx;
            double mass_l = -LoCy + LuCy;
            double mass_r = -RoCy + RuCy;
            double mass_oulr = (mass_o + mass_u + mass_l + mass_r) / 4;

            //===================== 
            // Masse z1 z2 z3 
            //=====================

            double mass_z2_o = (z2_O_Cx - z1_O_Cx) * 0.42447;
            double mass_z2_l = -(z2_L_Cy - z1_L_Cy) * 0.42447;
            double mass_z2_r = (z2_R_Cy - z1_R_Cy) * 0.42447;
            double mass_z2_u = -(z2_U_Cx - z1_U_Cx) * 0.42447;

            double mass_z3_o = (z3_O_Cx - z1_O_Cx) * 0.42447;
            double mass_z3_l = -(z3_L_Cy - z1_L_Cy) * 0.42447;
            double mass_z3_r = (z3_R_Cy - z1_R_Cy) * 0.42447;
            double mass_z3_u = -(z3_U_Cx - z1_U_Cx) * 0.42447;

            double mass_z4_o = (z4_O_Cx - z1_O_Cx) * 0.42447;
            double mass_z4_l = -(z4_L_Cy - z1_L_Cy) * 0.42447;
            double mass_z4_r = (z4_R_Cy - z1_R_Cy) * 0.42447;
            double mass_z4_u = -(z4_U_Cx - z1_U_Cx) * 0.42447;

            double mass_z2 = (mass_z2_o + mass_z2_l + mass_z2_r + mass_z2_u) / 4;
            double mass_z3 = (mass_z3_o + mass_z3_l + mass_z3_r + mass_z3_u) / 4;
            double mass_z4 = (mass_z4_o + mass_z4_l + mass_z4_r + mass_z4_u) / 4;


            //===================== 
            // Masse z_rand_2 z_rand_3 z_rand_4 
            //=====================

            double mass_z_rand_2 = (z_rand_2_Cz - z_rand_1_Cz);
            double mass_z_rand_3 = (z_rand_3_Cz - z_rand_1_Cz);
            double mass_z_rand_4 = (z_rand_4_Cz - z_rand_1_Cz);


            //===================== 
            // unif Center  
            //===================== 
            double unif_cent_Cx = unif_cent.CenterPoint.x;
            double unif_cent_Cy = unif_cent.CenterPoint.y;
            double unif_cent_Cz = unif_cent.CenterPoint.z;
            // Umrechnung der cm in VoxelIndex
            int unif_cent_Ix = Convert.ToInt32(Math.Ceiling((unif_cent.CenterPoint.x - oX) / resX));
            int unif_cent_Iy = Convert.ToInt32(Math.Ceiling((unif_cent.CenterPoint.y - oY) / resY));
            int unif_cent_Iz = Convert.ToInt32(Math.Ceiling((unif_cent.CenterPoint.z - oZ) / resZ));
            // HU
            double unif_cent_IV = MittelwertAuslesen(unif_cent_Ix, unif_cent_Iy, unif_cent_Iz, image);
            double unif_streu_cent_IV = StreuungAuslesen(unif_cent_Ix, unif_cent_Iy, unif_cent_Iz, image);

            //===================== 
            // unif oben  
            //===================== 
            double unif_ao_Cx = unif_ao.CenterPoint.x;
            double unif_ao_Cy = unif_ao.CenterPoint.y;
            double unif_ao_Cz = unif_ao.CenterPoint.z;

            double unif_io_Cx = unif_io.CenterPoint.x;
            double unif_io_Cy = unif_io.CenterPoint.y;
            double unif_io_Cz = unif_io.CenterPoint.z;
           
            // Umrechnung der cm in VoxelIndex
            int unif_ao_Ix = Convert.ToInt32(Math.Ceiling((unif_ao.CenterPoint.x - oX) / resX));
            int unif_ao_Iy = Convert.ToInt32(Math.Ceiling((unif_ao.CenterPoint.y - oY) / resY));
            int unif_ao_Iz = Convert.ToInt32(Math.Ceiling((unif_ao.CenterPoint.z - oZ) / resZ));
            // HU
            double unif_ao_IV = MittelwertAuslesen(unif_ao_Ix, unif_ao_Iy, unif_ao_Iz, image);
            double unif_streu_ao_IV = StreuungAuslesen(unif_ao_Ix, unif_ao_Iy, unif_ao_Iz, image);
            // Umrechnung der cm in VoxelIndex
            int unif_io_Ix = Convert.ToInt32(Math.Ceiling((unif_io.CenterPoint.x - oX) / resX));
            int unif_io_Iy = Convert.ToInt32(Math.Ceiling((unif_io.CenterPoint.y - oY) / resY));
            int unif_io_Iz = Convert.ToInt32(Math.Ceiling((unif_io.CenterPoint.z - oZ) / resZ));
            // HU
            double unif_io_IV = MittelwertAuslesen(unif_io_Ix, unif_io_Iy, unif_io_Iz, image);
            double unif_streu_io_IV = StreuungAuslesen(unif_io_Ix, unif_io_Iy, unif_io_Iz, image);

            //===================== 
            // unif rechts
            //===================== 
            double unif_ar_Cx = unif_ar.CenterPoint.x;
            double unif_ar_Cy = unif_ar.CenterPoint.y;
            double unif_ar_Cz = unif_ar.CenterPoint.z;

            double unif_ir_Cx = unif_ir.CenterPoint.x;
            double unif_ir_Cy = unif_ir.CenterPoint.y;
            double unif_ir_Cz = unif_ir.CenterPoint.z;

            // Umrechnung der cm in VoxelIndex
            int unif_ar_Ix = Convert.ToInt32(Math.Ceiling((unif_ar.CenterPoint.x - oX) / resX));
            int unif_ar_Iy = Convert.ToInt32(Math.Ceiling((unif_ar.CenterPoint.y - oY) / resY));
            int unif_ar_Iz = Convert.ToInt32(Math.Ceiling((unif_ar.CenterPoint.z - oZ) / resZ));
            // HU
            double unif_ar_IV = MittelwertAuslesen(unif_ar_Ix, unif_ar_Iy, unif_ar_Iz, image);
            double unif_streu_ar_IV = StreuungAuslesen(unif_ar_Ix, unif_ar_Iy, unif_ar_Iz, image);
            // Umrechnung der cm in VoxelIndex
            int unif_ir_Ix = Convert.ToInt32(Math.Ceiling((unif_ir.CenterPoint.x - oX) / resX));
            int unif_ir_Iy = Convert.ToInt32(Math.Ceiling((unif_ir.CenterPoint.y - oY) / resY));
            int unif_ir_Iz = Convert.ToInt32(Math.Ceiling((unif_ir.CenterPoint.z - oZ) / resZ));
            // HU
            double unif_ir_IV = MittelwertAuslesen(unif_ir_Ix, unif_ir_Iy, unif_ir_Iz, image);
            double unif_streu_ir_IV = StreuungAuslesen(unif_ir_Ix, unif_ir_Iy, unif_ir_Iz, image);

            //===================== 
            // unif unten
            //===================== 
            double unif_au_Cx = unif_au.CenterPoint.x;
            double unif_au_Cy = unif_au.CenterPoint.y;
            double unif_au_Cz = unif_au.CenterPoint.z;

            double unif_iu_Cx = unif_iu.CenterPoint.x;
            double unif_iu_Cy = unif_iu.CenterPoint.y;
            double unif_iu_Cz = unif_iu.CenterPoint.z;

            // Umrechnung der cm in VoxelIndex
            int unif_au_Ix = Convert.ToInt32(Math.Ceiling((unif_au.CenterPoint.x - oX) / resX));
            int unif_au_Iy = Convert.ToInt32(Math.Ceiling((unif_au.CenterPoint.y - oY) / resY));
            int unif_au_Iz = Convert.ToInt32(Math.Ceiling((unif_au.CenterPoint.z - oZ) / resZ));
            // HU
            double unif_au_IV = MittelwertAuslesen(unif_au_Ix, unif_au_Iy, unif_au_Iz, image);
            double unif_streu_au_IV = StreuungAuslesen(unif_au_Ix, unif_au_Iy, unif_au_Iz, image);
            // Umrechnung der cm in VoxelIndex
            int unif_iu_Ix = Convert.ToInt32(Math.Ceiling((unif_iu.CenterPoint.x - oX) / resX));
            int unif_iu_Iy = Convert.ToInt32(Math.Ceiling((unif_iu.CenterPoint.y - oY) / resY));
            int unif_iu_Iz = Convert.ToInt32(Math.Ceiling((unif_iu.CenterPoint.z - oZ) / resZ));
            // HU
            double unif_iu_IV = MittelwertAuslesen(unif_iu_Ix, unif_iu_Iy, unif_iu_Iz, image);
            double unif_streu_iu_IV = StreuungAuslesen(unif_iu_Ix, unif_iu_Iy, unif_iu_Iz, image);

            //===================== 
            // unif links
            //===================== 
            double unif_al_Cx = unif_al.CenterPoint.x;
            double unif_al_Cy = unif_al.CenterPoint.y;
            double unif_al_Cz = unif_al.CenterPoint.z;

            double unif_il_Cx = unif_il.CenterPoint.x;
            double unif_il_Cy = unif_il.CenterPoint.y;
            double unif_il_Cz = unif_il.CenterPoint.z;

            // Umrechnung der cm in VoxelIndex
            int unif_al_Ix = Convert.ToInt32(Math.Ceiling((unif_al.CenterPoint.x - oX) / resX));
            int unif_al_Iy = Convert.ToInt32(Math.Ceiling((unif_al.CenterPoint.y - oY) / resY));
            int unif_al_Iz = Convert.ToInt32(Math.Ceiling((unif_al.CenterPoint.z - oZ) / resZ));
            // HU
            double unif_al_IV = MittelwertAuslesen(unif_al_Ix, unif_al_Iy, unif_al_Iz, image);
            double unif_streu_al_IV = StreuungAuslesen(unif_al_Ix, unif_al_Iy, unif_al_Iz, image);
            // Umrechnung der cm in VoxelIndex
            int unif_il_Ix = Convert.ToInt32(Math.Ceiling((unif_il.CenterPoint.x - oX) / resX));
            int unif_il_Iy = Convert.ToInt32(Math.Ceiling((unif_il.CenterPoint.y - oY) / resY));
            int unif_il_Iz = Convert.ToInt32(Math.Ceiling((unif_il.CenterPoint.z - oZ) / resZ));
            // HU
            double unif_il_IV = MittelwertAuslesen(unif_il_Ix, unif_il_Iy, unif_il_Iz, image);
            double unif_streu_il_IV = StreuungAuslesen(unif_il_Ix, unif_il_Iy, unif_iu_Iz, image);


            //===================== 
            // unif sym 
            //===================== 
            double unif_symXa = 2 * (unif_ar_IV - unif_al_IV) / (unif_ar_IV + unif_al_IV + 2000) * 100;
            double unif_symXi = 2 * (unif_ir_IV - unif_il_IV) / (unif_ir_IV + unif_il_IV + 2000) * 100;
            double unif_symYa = 2 * (unif_ao_IV - unif_au_IV) / (unif_ao_IV + unif_au_IV + 2000) * 100;
            double unif_symYi = 2 * (unif_io_IV - unif_iu_IV) / (unif_io_IV + unif_iu_IV + 2000) * 100;

            //===================== 
            // unif hom 
            //===================== 
            double unif_homA  = 5 * (unif_ao_IV + unif_ar_IV + unif_au_IV + unif_al_IV - 4 * unif_cent_IV) / 4 / (unif_ao_IV + unif_ar_IV + unif_au_IV + unif_al_IV + unif_cent_IV + 5000) * 100;
            double unif_homI  = 5 * (unif_io_IV + unif_ir_IV + unif_iu_IV + unif_il_IV - 4 * unif_cent_IV) / 4 / (unif_io_IV + unif_ir_IV + unif_iu_IV + unif_il_IV + unif_cent_IV + 5000) * 100;
            double unif_homAI = 8 * (unif_ao_IV + unif_ar_IV + unif_au_IV + unif_al_IV - unif_io_IV - unif_ir_IV - unif_iu_IV - unif_il_IV) / 4 / (unif_ao_IV + unif_ar_IV + unif_au_IV + unif_al_IV + unif_io_IV + unif_ir_IV + unif_iu_IV + unif_il_IV + 8000) * 100;

            double[] streu = new double[9];
            streu[0] = unif_cent_IV;
            streu[1] = unif_ao_IV;
            streu[2] = unif_io_IV;
            streu[3] = unif_ar_IV;
            streu[4] = unif_ir_IV;
            streu[5] = unif_au_IV;
            streu[6] = unif_iu_IV;
            streu[7] = unif_al_IV;
            streu[8] = unif_il_IV;

            double unif_streu = (streu.Max() - streu.Min()) / streu.Average();


            //============================================
            // Berechnen der MTF-Parameter
            //============================================

            double radius = 47;
            // double faktorW = 180 / 3.14159265359;
            double MTFcX = mtf_0.CenterPoint.x;
            double MTFcY = mtf_0.CenterPoint.y;
            double MTFcZ = mtf_0.CenterPoint.z;

            int Anf = 36;
            int Dimens = 32;
            int Dimension = 100;

            double MTF = 0;

            double EE01min = 1;
            double EE01max = 1;
            
            double EE02min = 1;
            double EE02max = 1;

            double EE03min = 1;
            double EE03max = 1;

            double EE04min = 1;
            double EE04max = 1;

            double EE05min = 1;
            double EE05max = 1;

            double EE06min = 1;
            double EE06max = 1;

            double EE07min = 1;
            double EE07max = 1;

            double EE08min = 1;
            double EE08max = 1;

            double EE09min = 1;
            double EE09max = 1;

            double EE10min = 1;
            double EE10max = 1;

            double EE11min = 1;
            double EE11max = 1;

            // MTF01
            double MTF01startW = 194;
            double MTF01stopW = 219;
            double start01X = 0;
            double start01Y = 0;
            double stop01X = 0;
            double stop01Y = 0;
            int[] MTF01kurvI = new int[100];
            double[] MTF01kurvE = new double[100];
            MTFkurvE(ref MTF01kurvE, ref MTF01kurvI, Dimension, MTF01startW, MTF01stopW, radius, MTFcX, MTFcY, MTFcZ, oX, oY, oZ, resX, resY, resZ, ref start01X, ref start01Y, ref stop01X, ref stop01Y, image);
            double[] MTF01kurvEE = new double[Dimens];
            for (int i = 0; i < Dimens; i++)
            {
                MTF01kurvEE[i] = MTF01kurvE[Anf + i];
            }
            EE01min = MTF01kurvEE.Min();
            EE01max = MTF01kurvEE.Max();
            double MTF01 = Math.Abs(EE01max / EE01min);

            // MTF02
            double MTF02startW = 219;
            double MTF02stopW = 243;
            double start02X = 0;
            double start02Y = 0;
            double stop02X = 0;
            double stop02Y = 0;
            int[] MTF02kurvI = new int[100];
            double[] MTF02kurvE = new double[100];
            MTFkurvE(ref MTF02kurvE, ref MTF02kurvI, Dimension, MTF02startW, MTF02stopW, radius, MTFcX, MTFcY, MTFcZ, oX, oY, oZ, resX, resY, resZ, ref start02X, ref start02Y, ref stop02X, ref stop02Y, image);
            double[] MTF02kurvEE = new double[Dimens];
            for (int i = 0; i < Dimens; i++)
            {
                MTF02kurvEE[i] = MTF02kurvE[Anf + i];
            }
            EE02min = MTF02kurvEE.Min();
            EE02max = MTF02kurvEE.Max();
            double MTF02 = Math.Abs(EE02max / EE02min);

            // MTF03
            double MTF03startW = 243;
            double MTF03stopW = 265;
            double start03X = 0;
            double start03Y = 0;
            double stop03X = 0;
            double stop03Y = 0;
            int[] MTF03kurvI = new int[100];
            double[] MTF03kurvE = new double[100];
            MTFkurvE(ref MTF03kurvE, ref MTF03kurvI, Dimension, MTF03startW, MTF03stopW, radius, MTFcX, MTFcY, MTFcZ, oX, oY, oZ, resX, resY, resZ, ref start03X, ref start03Y, ref stop03X, ref stop03Y, image);
            double[] MTF03kurvEE = new double[Dimens];
            for (int i = 0; i < Dimens; i++)
            {
                MTF03kurvEE[i] = MTF03kurvE[Anf + i];
            }
            EE03min = MTF03kurvEE.Min();
            EE03max = MTF03kurvEE.Max();
            double MTF03 = Math.Abs(EE03max / EE03min);

            // MTF04
            double MTF04startW = 265;
            double MTF04stopW = 283;
            double start04X = 0;
            double start04Y = 0;
            double stop04X = 0;
            double stop04Y = 0;
            int[] MTF04kurvI = new int[100];
            double[] MTF04kurvE = new double[100];
            MTFkurvE(ref MTF04kurvE, ref MTF04kurvI, Dimension, MTF04startW, MTF04stopW, radius, MTFcX, MTFcY, MTFcZ, oX, oY, oZ, resX, resY, resZ, ref start04X, ref start04Y, ref stop04X, ref stop04Y, image);
            double[] MTF04kurvEE = new double[Dimens];
            for (int i = 0; i < Dimens; i++)
            {
                MTF04kurvEE[i] = MTF04kurvE[Anf + i];
            }
            EE04min = MTF04kurvEE.Min();
            EE04max = MTF04kurvEE.Max();
            double MTF04 = Math.Abs(EE04max / EE04min);

            // MTF05
            double MTF05startW = 283;
            double MTF05stopW = 301;
            double start05X = 0;
            double start05Y = 0;
            double stop05X = 0;
            double stop05Y = 0;
            int[] MTF05kurvI = new int[100];
            double[] MTF05kurvE = new double[100];
            MTFkurvE(ref MTF05kurvE, ref MTF05kurvI, Dimension, MTF05startW, MTF05stopW, radius, MTFcX, MTFcY, MTFcZ, oX, oY, oZ, resX, resY, resZ, ref start05X, ref start05Y, ref stop05X, ref stop05Y, image);
            double[] MTF05kurvEE = new double[Dimens];
            for (int i = 0; i < Dimens; i++)
            {
                MTF05kurvEE[i] = MTF05kurvE[Anf + i];
            }
            EE05min = MTF05kurvEE.Min();
            EE05max = MTF05kurvEE.Max();
            double MTF05 = Math.Abs(EE05max / EE05min);

            // MTF06
            double MTF06startW = 301;
            double MTF06stopW = 319;
            double start06X = 0;
            double start06Y = 0;
            double stop06X = 0;
            double stop06Y = 0;
            int[] MTF06kurvI = new int[100];
            double[] MTF06kurvE = new double[100];
            MTFkurvE(ref MTF06kurvE, ref MTF06kurvI, Dimension, MTF06startW, MTF06stopW, radius, MTFcX, MTFcY, MTFcZ, oX, oY, oZ, resX, resY, resZ, ref start06X, ref start06Y, ref stop06X, ref stop06Y, image);
            double[] MTF06kurvEE = new double[Dimens];
            for (int i = 0; i < Dimens; i++)
            {
                MTF06kurvEE[i] = MTF06kurvE[Anf + i];
            }
            EE06min = MTF06kurvEE.Min();
            EE06max = MTF06kurvEE.Max();
            double MTF06 = Math.Abs(EE06max / EE06min);

            // MTF07
            double MTF07startW = 319;
            double MTF07stopW = 335;
            double start07X = 0;
            double start07Y = 0;
            double stop07X = 0;
            double stop07Y = 0;
            int[] MTF07kurvI = new int[100];
            double[] MTF07kurvE = new double[100];
            MTFkurvE(ref MTF07kurvE, ref MTF07kurvI, Dimension, MTF07startW, MTF07stopW, radius, MTFcX, MTFcY, MTFcZ, oX, oY, oZ, resX, resY, resZ, ref start07X, ref start07Y, ref stop07X, ref stop07Y, image);
            double[] MTF07kurvEE = new double[Dimens];
            for (int i = 0; i < Dimens; i++)
            {
                MTF07kurvEE[i] = MTF07kurvE[Anf + i];
            }
            EE07min = MTF07kurvEE.Min();
            EE07max = MTF07kurvEE.Max();
            double MTF07 = Math.Abs(EE07max / EE07min);

            // MTF08
            double MTF08startW = 335;
            double MTF08stopW = 351;
            double start08X = 0;
            double start08Y = 0;
            double stop08X = 0;
            double stop08Y = 0;
            int[] MTF08kurvI = new int[100];
            double[] MTF08kurvE = new double[100];
            MTFkurvE(ref MTF08kurvE, ref MTF08kurvI, Dimension, MTF08startW, MTF08stopW, radius, MTFcX, MTFcY, MTFcZ, oX, oY, oZ, resX, resY, resZ, ref start08X, ref start08Y, ref stop08X, ref stop08Y, image);
            double[] MTF08kurvEE = new double[Dimens];
            for (int i = 0; i < Dimens; i++)
            {
                MTF08kurvEE[i] = MTF08kurvE[Anf + i];
            }
            EE08min = MTF08kurvEE.Min();
            EE08max = MTF08kurvEE.Max();
            double MTF08 = Math.Abs(EE08max / EE08min);

            // MTF09
            double MTF09startW = 351;
            double MTF09stopW = 6;
            double start09X = 0;
            double start09Y = 0;
            double stop09X = 0;
            double stop09Y = 0;
            int[] MTF09kurvI = new int[100];
            double[] MTF09kurvE = new double[100];
            MTFkurvE(ref MTF09kurvE, ref MTF09kurvI, Dimension, MTF09startW, MTF09stopW, radius, MTFcX, MTFcY, MTFcZ, oX, oY, oZ, resX, resY, resZ, ref start09X, ref start09Y, ref stop09X, ref stop09Y, image);
            double[] MTF09kurvEE = new double[Dimens];
            for (int i = 0; i < Dimens; i++)
            {
                MTF09kurvEE[i] = MTF09kurvE[Anf + i];
            }
            EE09min = MTF09kurvEE.Min();
            EE09max = MTF09kurvEE.Max();
            double MTF09 = Math.Abs(EE09max / EE09min);

            // MTF10
            double MTF10startW = 6;
            double MTF10stopW = 20;
            double start10X = 0;
            double start10Y = 0;
            double stop10X = 0;
            double stop10Y = 0;
            int[] MTF10kurvI = new int[100];
            double[] MTF10kurvE = new double[100];
            MTFkurvE(ref MTF10kurvE, ref MTF10kurvI, Dimension, MTF10startW, MTF10stopW, radius, MTFcX, MTFcY, MTFcZ, oX, oY, oZ, resX, resY, resZ, ref start10X, ref start10Y, ref stop10X, ref stop10Y, image);
            double[] MTF10kurvEE = new double[Dimens];
            for (int i = 0; i < Dimens; i++)
            {
                MTF10kurvEE[i] = MTF10kurvE[Anf + i];
            }
            EE10min = MTF10kurvEE.Min();
            EE10max = MTF10kurvEE.Max();
            double MTF10 = Math.Abs(EE10max / EE10min);

            // MTF11
            double MTF11startW = 20;
            double MTF11stopW = 35;
            double start11X = 0;
            double start11Y = 0;
            double stop11X = 0;
            double stop11Y = 0;
            int[] MTF11kurvI = new int[100];
            double[] MTF11kurvE = new double[100];
            MTFkurvE(ref MTF11kurvE, ref MTF11kurvI, Dimension, MTF11startW, MTF11stopW, radius, MTFcX, MTFcY, MTFcZ, oX, oY, oZ, resX, resY, resZ, ref start11X, ref start11Y, ref stop11X, ref stop11Y, image);
            double[] MTF11kurvEE = new double[Dimens];
            for (int i = 0; i < Dimens; i++)
            {
                MTF11kurvEE[i] = MTF11kurvE[Anf + i];
            }
            EE11min = MTF11kurvEE.Min();
            EE11max = MTF11kurvEE.Max();
            double MTF11 = Math.Abs(EE11max / EE11min);

            int[] MTFkurvII = new int[Dimens];

            

           

            double Schwelle = 1.3;

            if (MTF01 > Schwelle)
            {
                MTF = 1;
            }
            if (MTF02 > Schwelle)
            {
                MTF = 2;
            }
            if (MTF03 > Schwelle)
            {
                MTF = 3;
            }
            if (MTF04 > Schwelle)
            {
                MTF = 4;
            }
            if (MTF05 > Schwelle)
            {
                MTF = 5;
            }
            if (MTF06 > Schwelle)
            {
                MTF = 6;
            }
            if (MTF07 > Schwelle)
            {
                MTF = 7;
            }
            // if (MTF08 > Schwelle)
            // {
            //     MTF = 8;
            // }
            // if (MTF09 > Schwelle)
            // {
            //     MTF = 9;
            // }
            // if (MTF10 > Schwelle)
            // {
            //     MTF = 10;
            // }
            // if (MTF11 > Schwelle)
            // {
            //     MTF = 11;
            // }

            //============================================
            // Berechnen der Uniformität
            //============================================







            //============================================
            // Ausdruck der Werte in einem Message Fenster
            //============================================
            string message00 = "------------------------------------------------------------------------";
            string messageHU = "HU-Werte";
            string message = string.Format("Air       ", airCx, airCy, airCz, "\n\r", "Acrylic   ", acrylicCx, acrylicCy, acrylicCz);
            string message0 = string.Format("userOrign   " + "\t" + uoX.ToString("F1") + " " + uoY.ToString("F1") + " " + uoZ.ToString("F1"));
            string message1 = string.Format("Air               " + "\t" + airCx.ToString("F1") + " " + airCy.ToString("F1") + " " + airCz.ToString("F1") + "\t\t" + airIV.ToString("F2"));
            string message2 = string.Format("Teflon        " + "\t" + teflonCx.ToString("F1") + " " + teflonCy.ToString("F1") + " " + teflonCz.ToString("F1") + "\t\t" + teflonIV.ToString("F2"));
            string message3 = string.Format("Deldrine    " + "\t" + deldrineCx.ToString("F1") + " " + deldrineCy.ToString("F1") + " " + deldrineCz.ToString("F1") + "\t\t" + deldrineIV.ToString("F2"));
            string message4 = string.Format("Acrylic       " + "\t" + acrylicCx.ToString("F1") + " " + acrylicCy.ToString("F1") + " " + acrylicCz.ToString("F1") + "\t\t" + acrylicIV.ToString("F2"));
            string message5 = string.Format("PolyStyrene " + "\t" + polystyreneCx.ToString("F1") + " " + polystyreneCy.ToString("F1") + " " + polystyreneCz.ToString("F1") + "\t\t" + polystyreneIV.ToString("F2"));
            string message6 = string.Format("LDPE          " + "\t" + ldpeCx.ToString("F1") + " " + ldpeCy.ToString("F1") + " " + ldpeCz.ToString("F1") + "\t\t" + ldpeIV.ToString("F2"));
            string message7 = string.Format("PMP           " + "\t" + pmpCx.ToString("F1") + " " + pmpCy.ToString("F1") + " " + pmpCz.ToString("F1") + "\t\t" + pmpIV.ToString("F2"));

            string messageXY = "Transversal-Masse";
            string message21 = string.Format("oben           " + "\t" + mass_o.ToString("F2"));
            string message22 = string.Format("links             " + "\t" + mass_l.ToString("F2"));
            string message23 = string.Format("rechts         " + "\t" + mass_r.ToString("F2"));
            string message24 = string.Format("unten          " + "\t" + mass_u.ToString("F2"));
            string message20MW = string.Format("Mittelwert       " + "\t" + mass_oulr.ToString("F2"));

            string messageZ = "Longitudinal-Masse";
            string message31 = string.Format("Z oben        " + "\t" + mass_z2_o.ToString("F2") + "\t" + mass_z3_o.ToString("F2") + "\t" + mass_z4_o.ToString("F2"));
            string message32 = string.Format("Z links           " + "\t" + mass_z2_l.ToString("F2") + "\t" + mass_z3_l.ToString("F2") + "\t" + mass_z4_l.ToString("F2"));
            string message33 = string.Format("Z rechts      " + "\t" + mass_z2_r.ToString("F2") + "\t" + mass_z3_r.ToString("F2") + "\t" + mass_z4_r.ToString("F2"));
            string message34 = string.Format("Z unten       " + "\t" + mass_z2_u.ToString("F2") + "\t" + mass_z3_u.ToString("F2") + "\t" + mass_z4_u.ToString("F2"));
            string message30MW = string.Format("Mittelwert       " + "\t" + mass_z2.ToString("F2") + "\t" + mass_z3.ToString("F2") + "\t" + mass_z4.ToString("F2"));
            string messageMTF = string.Format("MTF       " + "\t" + MTF.ToString() + "\t" + "MTF05    " + "\t" + MTF05.ToString("F2"));


            // normale MessageBox
            MessageBox.Show(messageHU + "\r\n" + message00 + "\r\n" + message1 + "\r\n" + message2 + "\r\n" + message3 + "\r\n" + message4
                + "\r\n" + message5 + "\r\n" + message6 + "\r\n" + message7 + "\r\n"
                + "\r\n" + messageXY + "\r\n" + message00 + "\r\n" + message21 + "\r\n" + message22 + "\r\n" + message23 + "\r\n" + message24 + "\r\n" + message20MW + "\r\n"
                + "\r\n" + messageZ + "\r\n" + message00 + "\r\n" + message31 + "\r\n" + message32 + "\r\n" + message33 + "\r\n" + message34 + "\r\n"
                + "\r\n" + message30MW + "\r\n" + "\r\n" + messageMTF 
                + "\r\n" + "\r\n" + "oben"
                + "\r\n" + unif_ao_IV.ToString("F2") + "\t" + unif_io_IV.ToString("F2")
                + "\r\n" + "\r\n" + "rechts"
                + "\r\n" + unif_ar_IV.ToString("F2") + "\t" + unif_ir_IV.ToString("F2")
                + "\r\n" + "\r\n" + "unten"
                + "\r\n" + unif_au_IV.ToString("F2") + "\t" + unif_iu_IV.ToString("F2")
                + "\r\n" + "\r\n" + "links"
                + "\r\n" + unif_al_IV.ToString("F2") + "\t" + unif_il_IV.ToString("F2"),
                SCRIPT_NAME, MessageBoxButton.YesNoCancel, MessageBoxImage.Exclamation); 


           


            //===================
            // in Excel schreiben
            //===================
            
            // Normales Abspeichern der TestErgebnisse
            UpdateExcel_HU(Datum, Zeit, "HU_Daten", airIV.ToString("F2"), teflonIV.ToString("F2"), deldrineIV.ToString("F2"),
                acrylicIV.ToString("F2"), polystyreneIV.ToString("F2"), ldpeIV.ToString("F2"), pmpIV.ToString("F2"),
                mass_oulr.ToString("F2"), mass_z2.ToString("F2"), mass_z3.ToString("F2"), mass_z4.ToString("F2"),
                mass_z_rand_2.ToString("F2"), mass_z_rand_3.ToString("F2"), mass_z_rand_4.ToString("F2"),
                MTF.ToString(), MTF05.ToString("F2"));

            MessageBox.Show("HU Sheet geschrieben", SCRIPT_NAME, MessageBoxButton.YesNoCancel, MessageBoxImage.Exclamation);

            UpdateExcel_unif(Datum, Zeit, "unif_Daten", 
                unif_symXa.ToString("F2"), unif_symXi.ToString("F2"), unif_symYa.ToString("F2"), unif_symYi.ToString("F2"),
                unif_homA.ToString("F2"), unif_homI.ToString("F2"), unif_homAI.ToString("F2"), unif_streu.ToString("F2"), 
                unif_streu_cent_IV.ToString("F2"), 
                unif_streu_ao_IV.ToString("F2"), unif_streu_ar_IV.ToString("F2"), unif_streu_au_IV.ToString("F2"), unif_streu_al_IV.ToString("F2"), 
                unif_streu_io_IV.ToString("F2"), unif_streu_ir_IV.ToString("F2"), unif_streu_iu_IV.ToString("F2"), unif_streu_il_IV.ToString("F2"));

            MessageBox.Show("Unif Sheet geschrieben", SCRIPT_NAME, MessageBoxButton.YesNoCancel, MessageBoxImage.Exclamation);

            //++++++++++++++++++++++++++++++++++++++++++++++++++

            // window.Title = "MTF-KurvenPlot";

        }


        //##########################################################################################################
        //++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        void MTFkurvE(ref double[] MTFkurv0E, ref int[] MTFkurv0I, int Dimension,
            double MTFstart0W, double MTFstop0W, double radius0,
            double MTFc0X, double MTFc0Y, double MTFc0Z,
            double o0X, double o0Y, double o0Z,
            double res0X, double res0Y, double res0Z,
            ref double start0X, ref double start0Y,
            ref double stop0X, ref double stop0Y,
            VMS.TPS.Common.Model.API.Image Bild)
        {

            double faktor0W = 3.14159265359 / 180;
            // Indizes der MTF Funktionen
            int[] MTFkurvI0X = new int[100];
            int[] MTFkurvI0Y = new int[100];
            int[] MTFkurvI0Z = new int[100];
            double[] MTFkurv0W = new double[100];
            // xy Koordinaten der MTF Funktion
            double[] MTFkurv0X = new double[100];
            double[] MTFkurv0Y = new double[100];
            double[] MTFkurv0Z = new double[100];

            double step0W = (MTFstop0W - MTFstart0W) / (Dimension - 1);

            // Winkelschritte
            int[,] bildPlane = new int[Bild.XSize, Bild.YSize];

            for (int i = 0; i < Dimension; i++)
            {
                // Bestimmen der xy Koordinaten für MTF Linienpaare
                MTFkurv0W[i] = MTFstart0W + step0W * i;
                MTFkurv0X[i] = MTFc0X + radius0 * Math.Cos(faktor0W * MTFkurv0W[i]);
                MTFkurv0Y[i] = MTFc0Y - radius0 * Math.Sin(faktor0W * MTFkurv0W[i]);
                MTFkurv0Z[i] = MTFc0Z;
                // StartPosition der MTF
                if (i == 0)
                {
                    start0X = MTFkurv0X[i];
                    start0Y = MTFkurv0Y[i];
                }
                // Stopposition der MTF
                if (i == Dimension - 1)
                {
                    stop0X = MTFkurv0X[i];
                    stop0Y = MTFkurv0Y[i];
                }
                // Iniizes der MTF Funktionen
                MTFkurvI0X[i] = (int)Math.Round((MTFkurv0X[i] - o0X) / res0X);
                MTFkurvI0Y[i] = (int)Math.Round((MTFkurv0Y[i] - o0Y) / res0Y);
                MTFkurvI0Z[i] = (int)Math.Round((MTFkurv0Z[i] - o0Z) / res0Z);

                // Auslesen der HU Werte bei MTF
   // Fehler !!!!     muss nach vorne !!!
                if (i == 0)
                {
                    Bild.GetVoxels(MTFkurvI0Z[i], bildPlane);
                }
                MTFkurv0E[i] = (Bild.VoxelToDisplayValue(bildPlane[MTFkurvI0X[i], MTFkurvI0Y[i]]));
                MTFkurv0I[i] = i;

            }



        }


        //#########################################################################################

        //#########################################################################################
        // normale Excelausgabe der Testergebnisse
        private void UpdateExcel_HU(string Datum, string Zeit, string sheetName,
                string airIV, string teflonIV, string deldrineIV,
                string acrylicIV, string polystyreneIV, string ldpeIV, string pmpIV,
                string mass_oulr, string mass_z2, string mass_z3, string mass_z4,
                string mass_z_rand_2, string mass_z_rand_3, string mass_z_rand_4,
                string MTF, string MTF05)
        {

            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel._Worksheet oSheet = null;

            try
            {
                // normal abspeichern in Excwl

                oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Open("Q:/CT/QA_xls/CT_HUCheckCadphan.xlsx");
                oSheet = String.IsNullOrEmpty(sheetName) ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets[sheetName];

                string sReihe = oSheet.Cells[3, 2].Value == null ? "-" : oSheet.Cells[3, 2].Value.ToString();
                int iReihe = Convert.ToInt32(sReihe);
                iReihe = iReihe + 1;

                oSheet.Cells[iReihe, 1] = Datum;
                oSheet.Cells[iReihe, 2] = Zeit;

                oSheet.Cells[iReihe, 3] = airIV;
                oSheet.Cells[iReihe, 4] = teflonIV;
                oSheet.Cells[iReihe, 5] = deldrineIV;
                oSheet.Cells[iReihe, 6] = acrylicIV;
                oSheet.Cells[iReihe, 7] = polystyreneIV;
                oSheet.Cells[iReihe, 8] = ldpeIV;
                oSheet.Cells[iReihe, 9] = pmpIV;

                oSheet.Cells[iReihe, 11] = mass_oulr;

                oSheet.Cells[iReihe, 13] = mass_z2;
                oSheet.Cells[iReihe, 14] = mass_z3;
                oSheet.Cells[iReihe, 15] = mass_z4;

                oSheet.Cells[iReihe, 17] = mass_z_rand_2;
                oSheet.Cells[iReihe, 18] = mass_z_rand_3;
                oSheet.Cells[iReihe, 19] = mass_z_rand_4;
           
                oSheet.Cells[iReihe, 21] = MTF;
                oSheet.Cells[iReihe, 22] = MTF05;

                sReihe = Convert.ToString(iReihe);

                oSheet.Cells[3, 2] = sReihe;


                oWB.Save();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close(true, null, null);
                    oXL.Quit();
                }

                //oWB.Close(false);

            }

            // MessageBox.Show("Done");

        }
        //.........................................................................................
        private void UpdateExcel_unif(string Datum, string Zeit, string sheetName,
                string unif_symXa, string unif_symXi, string unif_symYa, string unif_symYi,
                string unif_homA, string unif_homI, string unif_homAI, string unif_streu, 
                string unif_streu_cent_IV, 
                string unif_streu_ao_IV, string unif_streu_ar_IV, string unif_streu_au_IV, string unif_streu_al_IV, 
                string unif_streu_io_IV, string unif_streu_ir_IV, string unif_streu_iu_IV, string unif_streu_il_IV)
        {

            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel._Worksheet oSheet = null;


            try
            {
                // normal abspeichern in Excwl

                oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Open("Q:/CT/QA_xls/CT_HUCheckCadphan.xlsx");
                oSheet = String.IsNullOrEmpty(sheetName) ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets[sheetName];

                string sReihe = oSheet.Cells[3, 2].Value == null ? "-" : oSheet.Cells[3, 2].Value.ToString();
                int iReihe = Convert.ToInt32(sReihe);
                iReihe = iReihe + 1;

                oSheet.Cells[iReihe, 1] = Datum;
                oSheet.Cells[iReihe, 2] = Zeit;

                oSheet.Cells[iReihe, 3] = unif_symXa;
                oSheet.Cells[iReihe, 4] = unif_symXi;
                oSheet.Cells[iReihe, 5] = unif_symYa;
                oSheet.Cells[iReihe, 6] = unif_symYi;

                oSheet.Cells[iReihe, 8] = unif_homA;
                oSheet.Cells[iReihe, 9] = unif_homI;
                oSheet.Cells[iReihe, 10] = unif_homAI;
                oSheet.Cells[iReihe, 11] = unif_streu;

                oSheet.Cells[iReihe, 13] = unif_streu_cent_IV;
                oSheet.Cells[iReihe, 14] = unif_streu_ao_IV;
                oSheet.Cells[iReihe, 15] = unif_streu_ar_IV;
                oSheet.Cells[iReihe, 16] = unif_streu_au_IV;
                oSheet.Cells[iReihe, 17] = unif_streu_al_IV;
                oSheet.Cells[iReihe, 18] = unif_streu_io_IV;
                oSheet.Cells[iReihe, 19] = unif_streu_ir_IV;
                oSheet.Cells[iReihe, 20] = unif_streu_iu_IV; 
                oSheet.Cells[iReihe, 21] = unif_streu_il_IV;
                               
                sReihe = Convert.ToString(iReihe);

                oSheet.Cells[3, 2] = sReihe;


                oWB.Save();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close(true, null, null);
                    oXL.Quit();
                }

                //oWB.Close(false);

            }

            // MessageBox.Show("Done");

        }

        // ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        // Excelausgabe zum Debuggen
        private void UpdateExcel1(string Datum, string Zeit, string sheetName, 
            double uoX, double uoY, double uoZ, 
            double oX, double oY, double oZ, 
            double sizeX, double sizeY, double sizeZ, 
            double resX, double resY, double resZ, 
            double MTFcX, double MTFcY, double MTFcZ, 
            double[] Kurv1, double[] Kurv2, double[] Kurv3, double[] Kurv4, double[] Kurv5,
            double start01X, double start01Y, double stop01X, double stop01Y,
            double start02X, double start02Y, double stop02X, double stop02Y,
            double start03X, double start03Y, double stop03X, double stop03Y,
            double start04X, double start04Y, double stop04X, double stop04Y,
            double start05X, double start05Y, double stop05X, double stop05Y,
            double MTF01, double MTF02, double MTF03, double MTF04, double MTF05, 
            double MTF06, double MTF07, double MTF08, double MTF09, double MTF10, double MTF11)
        {

            Excel.Application oXL = null;
            Excel._Workbook oWB = null;
            Excel._Worksheet oSheet = null;


            try
            {
                // normal abspeichern in Excwl

                oXL = new Microsoft.Office.Interop.Excel.Application();
                oWB = oXL.Workbooks.Open("Q:/CT/QA_xls/CT_Diagramme.xlsx");
                oSheet = String.IsNullOrEmpty(sheetName) ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets[sheetName];

                // Diagramm abspeichern in Excel

                oSheet = String.IsNullOrEmpty("Chart1") ? (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet : (Microsoft.Office.Interop.Excel._Worksheet)oWB.Worksheets[sheetName];

                for (int i = 1; i < 32; i++)
                {
                    oSheet.Cells[i, 1] = i.ToString();
                    oSheet.Cells[i, 2] = Kurv1[i].ToString("F1");
                    oSheet.Cells[i, 3] = Kurv2[i].ToString("F1");
                    oSheet.Cells[i, 4] = Kurv3[i].ToString("F1");
                    oSheet.Cells[i, 5] = Kurv4[i].ToString("F1");
                    oSheet.Cells[i, 6] = Kurv5[i].ToString("F1");

                }
                // User Origin
                oSheet.Cells[2, 10] = uoX.ToString("F2");
                oSheet.Cells[3, 10] = uoY.ToString("F2");
                oSheet.Cells[4, 10] = uoZ.ToString("F2");
                // DICOM Origin
                oSheet.Cells[2, 11] = oX.ToString("F2");
                oSheet.Cells[3, 11] = oY.ToString("F2");
                oSheet.Cells[4, 11] = oZ.ToString("F2");
                // Index Max X Y Z
                oSheet.Cells[2, 12] = sizeX.ToString("F2");
                oSheet.Cells[3, 12] = sizeY.ToString("F2");
                oSheet.Cells[4, 12] = sizeZ.ToString("F2");
                // Resolution X Y Z
                oSheet.Cells[2, 13] = resX.ToString("F2");
                oSheet.Cells[3, 13] = resY.ToString("F2");
                oSheet.Cells[4, 13] = resZ.ToString("F2");
                // Schwerpunkt MTF Kreis
                oSheet.Cells[2, 14] = MTFcX.ToString("F2");
                oSheet.Cells[3, 14] = MTFcY.ToString("F2");
                oSheet.Cells[4, 14] = MTFcZ.ToString("F2");
                              
                // StartStop MTFkurv01
                oSheet.Cells[62, 2] = start01X.ToString("F2");
                oSheet.Cells[63, 2] = start01Y.ToString("F2");

                oSheet.Cells[66, 2] = stop01X.ToString("F2");
                oSheet.Cells[67, 2] = stop01Y.ToString("F2");

                // StartStop MTFkurv02
                oSheet.Cells[62, 3] = start02X.ToString("F2");
                oSheet.Cells[63, 3] = start02Y.ToString("F2");

                oSheet.Cells[66, 3] = stop02X.ToString("F2");
                oSheet.Cells[67, 3] = stop02Y.ToString("F2");

                // StartStop MTFkurv03
                oSheet.Cells[62, 4] = start03X.ToString("F2");
                oSheet.Cells[63, 4] = start03Y.ToString("F2");

                oSheet.Cells[66, 4] = stop03X.ToString("F2");
                oSheet.Cells[67, 4] = stop03Y.ToString("F2");

                // StartStop MTFkurv04
                oSheet.Cells[62, 5] = start04X.ToString("F2");
                oSheet.Cells[63, 5] = start04Y.ToString("F2");

                oSheet.Cells[66, 5] = stop04X.ToString("F2");
                oSheet.Cells[67, 5] = stop04Y.ToString("F2");

                // StartStop MTFkurv05
                oSheet.Cells[62, 6] = start05X.ToString("F2");
                oSheet.Cells[63, 6] = start05Y.ToString("F2");

                oSheet.Cells[66, 6] = stop05X.ToString("F2");
                oSheet.Cells[67, 6] = stop05Y.ToString("F2");

                // MTF Werte
                oSheet.Cells[2, 8] = MTF01.ToString("F2");
                oSheet.Cells[3, 8] = MTF02.ToString("F2");
                oSheet.Cells[4, 8] = MTF03.ToString("F2");
                oSheet.Cells[5, 8] = MTF04.ToString("F2");
                oSheet.Cells[6, 8] = MTF05.ToString("F2");
                oSheet.Cells[7, 8] = MTF06.ToString("F2");
                oSheet.Cells[8, 8] = MTF07.ToString("F2");
                oSheet.Cells[9, 8] = MTF08.ToString("F2");
                oSheet.Cells[10, 8] = MTF09.ToString("F2");
                oSheet.Cells[11, 8] = MTF10.ToString("F2");
                oSheet.Cells[12, 8] = MTF11.ToString("F2");


                oWB.Save();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                if (oWB != null)
                {
                    oWB.Close(true, null, null);
                    oXL.Quit();
                }

                //oWB.Close(false);

            }

            // MessageBox.Show("Done");

        }





    }
}
