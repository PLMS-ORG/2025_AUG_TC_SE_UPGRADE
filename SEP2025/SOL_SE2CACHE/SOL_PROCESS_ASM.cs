using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LTC_SE2CACHE
{
    class SOL_PROCESS_ASM
    {
        static String COMMERCIAL = "Commercial";
        private static string WELDED_ASSEMBLY = "Welded Assembly";

        public static void PROCESS_ASM(String StageDir, String outputFolderPath, String logFilePath)
        {
            string[] AsmFiles = Directory.GetFiles(StageDir, "*")
                                         .Select(path => Path.GetFullPath(path))
                                         .Where(x => (x.EndsWith(".asm") || x.EndsWith(".ASM")))
                                         .ToArray();
            if (AsmFiles == null || AsmFiles.Length == 0)
            {
                Console.WriteLine("SOL_PROCESS_ASM:- NO ASSEMBLY FILES IDENTIFIED");
                return;
            }

            foreach (String asmFILEFullPath in AsmFiles)
            {
                Utility.Log(asmFILEFullPath + " Translation Started..", logFilePath);

                // 1 AUG 2019 - Dont process ref files
                String asmFILE = Path.GetFileName(asmFILEFullPath);
                string asyFileName = Path.GetFileName(asmFILEFullPath);
                string itemIDofFile = LTC_SE_SET_PROPERTIES.getProperty("ASSEMBLY", asyFileName, "ITEM_ID");
                if (Program.part_Asm_SM_Weld_FilesToProcess.Count != 0)
                {
                    string asmFileWoExtn = Path.GetFileNameWithoutExtension(asmFILEFullPath);
                    Utility.Log("\n SOL_PROCESS_ASM:- " + asmFileWoExtn + "Checking if part belongs to ref category", logFilePath);
                    if (Program.part_Asm_SM_Weld_FilesToProcess.Contains(asmFileWoExtn.ToLower()) == false)
                    {
                        Utility.Log("SOL_PROCESS_ASM:- " + asmFileWoExtn + " belongs to ref category. Skipping", logFilePath);
                        Utility.Log("SOL_PROCESS_PART:- " + asmFileWoExtn + " Making sure part belongs to ref category with ID", logFilePath);
                        Utility.Log("SOL_PROCESS_PART:- " + "checking if " + itemIDofFile + " belongs to ref category", logFilePath);
                        if (Program.part_Asm_SM_Weld_FilesToProcess.Contains(itemIDofFile.ToLower()) == false)
                        {
                            Utility.Log("SOL_PROCESS_PART:- " + itemIDofFile + " belongs to ref category with ID. Skipping", logFilePath);
                            Program.totalFilesProcessed++;
                            continue;
                        }
                        else
                        {
                            Utility.Log("SOL_PROCESS_PART:- " + itemIDofFile + " does not belong to ref category with ID", logFilePath);
                        }
                    }
                    else
                        Utility.Log("SOL_PROCESS_ASM:- " + asmFileWoExtn + " does not belong to ref category.", logFilePath);
                }
                else
                    Utility.Log("\n SOL_PROCESS_DRAFT:- SEAssyOut.txt is not available ", logFilePath);

                // 11 OCT 2018 - Dont Process Asm which have Release Status Count > 0
                //String Release_Status_Count = LTC_SE_SET_PROPERTIES.getProperty("ASSEMBLY", asmFILE, "RELEASE_STATUS");
                //int Release_Count = 0;
                //try
                //{
                //    Release_Count = int.Parse(Release_Status_Count);
                //}
                //catch (Exception ex)
                //{
                //    Utility.Log("SOL_PROCESS_ASM:- Could Not Parse the Release Count..Exception: " + asmFILE, logFilePath);
                //    Utility.Log("Problem encountered while finding release status of " + itemIDofFile, Program.failureFilePath);
                //    Utility.Log("SOL_PROCESS_ASM:- Exception: " + ex.Message, logFilePath);
                //    Release_Count = 0;
                //}
                //if (Release_Count > 0 && Utility.IgnoreReleaseStatus.Equals("NO", StringComparison.OrdinalIgnoreCase) == true)
                //{
                //    Utility.Log("SOL_PROCESS_ASM:- Skip " + asmFILE + " ,Since File Has Release Status Already...", logFilePath);
                //    Program.totalFilesProcessed++;
                //    continue;
                //}


                //String partType = LTC_SE_SET_PROPERTIES.getProperty("ASSEMBLY", asmFILE, "PART_TYPE");

                //if (partType != null && partType.Equals("") == false && partType.Equals(COMMERCIAL, StringComparison.OrdinalIgnoreCase) == true)
                //{
                //    Utility.Log("SOL_PROCESS_ASM:- Skip PART_TYPE COMMERCIAL for " + asmFILE, logFilePath);
                //    Program.totalFilesProcessed++;
                //    continue;
                //}

                //String category = LTC_SE_SET_PROPERTIES.getProperty("ASSEMBLY", asmFILE, "CATEGORY");

                //if (category == null || category.Equals("") == true || category.Equals(WELDED_ASSEMBLY, StringComparison.OrdinalIgnoreCase) == false)
                //{
                //    Utility.Log("SOL_PROCESS_ASM:- Category is not Welded Assembly For: " + asmFILE, logFilePath);
                //    Program.totalFilesProcessed++;
                //    continue;

                //}

                //19-Oct-2019 If item id and rev id is passed ignore other files (for post release IR, all files need to be downloaded to stage directory)
                if (Program.revIDReceived != null & Program.itemIDRecieved != null)
                {
                    string itemIDFromFile = LTC_SE_SET_PROPERTIES.getProperty("ASSEMBLY", Path.GetFileName(asmFILEFullPath), "ITEM_ID");
                    Utility.Log("[ItemID] :" + itemIDFromFile, logFilePath);

                    string onlyItemToProcess = Program.itemIDRecieved;
                    Utility.Log("onlyItemToProcess :" + onlyItemToProcess, logFilePath);

                    if (itemIDFromFile.ToLower().Equals(onlyItemToProcess.ToLower()) == false)
                    {
                        Program.totalFilesProcessed++;
                        continue;
                    }
                }

                String Folder = outputFolderPath;
                //SOL_SE_TRANSLATE.SaveAssemblyAs(StageDir, asmFILE, "STP", logFilePath, Folder);
                SOL_SE_TRANSLATE assembly = new SOL_SE_TRANSLATE();
                assembly.SaveAssemblyAsSTATThread(asmFILEFullPath, "STP", logFilePath, Folder, itemIDofFile);

                Utility.Log(asmFILEFullPath + " Translation is completed...", logFilePath);
                Program.totalFilesProcessed++;
            }

        }
    }
}
