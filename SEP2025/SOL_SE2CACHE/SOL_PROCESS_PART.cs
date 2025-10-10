using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LTC_SE2CACHE
{
    class SOL_PROCESS_PART
    {
        private static string COMMERCIAL = "Commercial";
        private static string WELDED_ASSEMBLY = "Welded Assembly";

        public static void PROCESS_PART(String StageDir, String outputFolderPath, String logFilePath)
        {
            string[] PartFiles = Directory.GetFiles(StageDir, "*")
                                         .Select(path => Path.GetFullPath(path))
                                         .Where(x => (x.EndsWith(".par") || x.EndsWith(".PAR")))
                                         .ToArray();
            if (PartFiles == null || PartFiles.Length == 0)
            {
                Console.WriteLine("SOL_PROCESS_PART:- NO PART FILES IDENTIFIED");
                return;
            }

            foreach (String partFileFullPath in PartFiles)
            {
                Utility.Log(partFileFullPath + " Starting..Translation", logFilePath);

                // 1 AUG 2019 - Dont process ref files
                String partFile = Path.GetFileName(partFileFullPath);
                string partFileName = Path.GetFileName(partFileFullPath);
                string itemIDofFile = LTC_SE_SET_PROPERTIES.getProperty("PART", partFileName, "ITEM_ID");
                if (Program.part_Asm_SM_Weld_FilesToProcess.Count != 0)
                {
                    string partFileWoExtn = Path.GetFileNameWithoutExtension(partFileFullPath);
                    Utility.Log("\n SOL_PROCESS_PART:- " + partFileWoExtn + " Checking if part belongs to ref category", logFilePath);
                    if (Program.part_Asm_SM_Weld_FilesToProcess.Contains(partFileWoExtn.ToLower()) == false)
                    {
                        Utility.Log("SOL_PROCESS_PART:- " + partFileWoExtn + " belongs to ref category. Skipping", logFilePath);
                        Utility.Log("SOL_PROCESS_PART:- " + partFileWoExtn + " Making sure part belongs to ref category with ID", logFilePath);
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
                        Utility.Log("SOL_PROCESS_PART:- " + partFileWoExtn + " does not belong to ref category.", logFilePath);
                }
                else
                    Utility.Log("\n SOL_PROCESS_PART:- SEAssyOut.txt is not available ", logFilePath);

                //// 11 OCT 2018 - Dont Process PAR which have Release Status Count > 0
                //String Release_Status_Count = LTC_SE_SET_PROPERTIES.getProperty("PART", partFile, "RELEASE_STATUS");
                //int Release_Count = 0;
                //try
                //{
                //    Release_Count = int.Parse(Release_Status_Count);
                //}
                //catch (Exception ex)
                //{
                //    Utility.Log("SOL_PROCESS_PART:- Could Not Parse the Release Count..Exception: " + partFile, logFilePath);
                //    Utility.Log("Problem encountered while finding release status of " + itemIDofFile, Program.failureFilePath);
                //    Utility.Log("SOL_PROCESS_PART:- Exception: " + ex.Message, logFilePath);
                //    Release_Count = 0;
                //}
                //if (Release_Count > 0 && Utility.IgnoreReleaseStatus.Equals("NO", StringComparison.OrdinalIgnoreCase) == true)
                //{
                //    Utility.Log("SOL_PROCESS_PART:- Skip " + partFile + " ,Since File Has Release Status Already...", logFilePath);
                //    Program.totalFilesProcessed++;
                //    continue;
                //}

                //String partType = LTC_SE_SET_PROPERTIES.getProperty("PART", partFile, "PART_TYPE");
                //if (partType != null && partType.Equals("") == false && partType.Equals(COMMERCIAL, StringComparison.OrdinalIgnoreCase) == true)
                //{
                //    Utility.Log("SOL_PROCESS_PART:- Skip PART_TYPE COMMERCIAL for " + partFile, logFilePath);
                //    Program.totalFilesProcessed++;
                //    continue;

                //}

                //19-Oct-2019 If item id and rev id is passed ignore other files (for post release IR, all files need to be downloaded to stage directory)
                if (Program.revIDReceived != null & Program.itemIDRecieved != null)
                {
                    string itemIDFromFile = LTC_SE_SET_PROPERTIES.getProperty("PART", Path.GetFileName(partFileFullPath), "ITEM_ID");
                    Utility.Log("[ItemID] :" + itemIDFromFile, logFilePath);

                    string onlyItemToProcess = Program.itemIDRecieved;
                    Utility.Log("onlyItemToProcess :" + onlyItemToProcess, logFilePath);
                    if (itemIDFromFile.ToLower().Equals(onlyItemToProcess.ToLower()) == false)
                    {
                        Program.totalFilesProcessed++;
                        continue;
                    }
                }

                // 19-02-2019 Step File should be Generated for All Part Files.
                //String category = SOL_SE_SET_PROPERTIES.getProperty("PART", partFile, "CATEGORY");

                //if (category == null || category.Equals("") == true || category.Equals(WELDED_ASSEMBLY, StringComparison.OrdinalIgnoreCase) == false)
                //{
                //    Utility.Log("SOL_PROCESS_PART:- Category is not Welded Assembly For: " + partFile, logFilePath);
                //    continue;

                //}

                String Folder = outputFolderPath;
                //SOL_SE_TRANSLATE.SavePartAs(StageDir, partFile, "STP", logFilePath, Folder);
                SOL_SE_TRANSLATE part = new SOL_SE_TRANSLATE();
                part.SavePartAsSTATThread(StageDir, partFileFullPath, "STP", logFilePath, Folder, itemIDofFile);

                Utility.Log("-----------------------------------------------------------------: " + partFile, logFilePath);
                Program.totalFilesProcessed++;
            }


        }
    }
}
