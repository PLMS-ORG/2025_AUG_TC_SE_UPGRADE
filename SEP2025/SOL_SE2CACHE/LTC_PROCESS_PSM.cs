using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LTC_SE2CACHE
{
    class LTC_PROCESS_PSM
    {
        
        public static void PROCESS_SHEETMETAL_To_DXF(String StageDir, String outputFolderPath, String logFilePath)
        {
            string[] PartFiles = Directory.GetFiles(StageDir, "*")
                                             .Select(path => Path.GetFullPath(path))
                                             .Where(x => (x.EndsWith(".psm") || x.EndsWith(".PSM")))
                                             .ToArray();
            if (PartFiles == null || PartFiles.Length == 0)
            {
                Console.WriteLine("LTC_PROCESS_PSM:- NO SHEETMETAL FILES IDENTIFIED");
                return;
            }

            foreach (String psmFileFullPath in PartFiles)
            {
                Utility.Log(psmFileFullPath + " Translation Starting..", logFilePath);

                // 1 AUG 2019 - Dont process ref files
                String psmFile = Path.GetFileName(psmFileFullPath);
                string psmFileName = Path.GetFileName(psmFileFullPath);
                string itemIDofFile = LTC_SE_SET_PROPERTIES.getProperty("SHEETMETAL", psmFileName, "ITEM_ID");
                if (Program.part_Asm_SM_Weld_FilesToProcess.Count != 0)
                {
                    string partFileWoExtn = Path.GetFileNameWithoutExtension(psmFileFullPath);
                    Utility.Log("\n LTC_PROCESS_PSM:- " + partFileWoExtn + "Checking if part belongs to ref category", logFilePath);
                    if (Program.part_Asm_SM_Weld_FilesToProcess.Contains(partFileWoExtn.ToLower()) == false)
                    {
                        Utility.Log("LTC_PROCESS_PSM:- " + partFileWoExtn + " belongs to ref category. Skipping", logFilePath);
                        Utility.Log("LTC_PROCESS_PSM:- " + partFileWoExtn + " Making sure part belongs to ref category with ID", logFilePath);
                        Utility.Log("LTC_PROCESS_PSM:- " + "checking if " + itemIDofFile + " belongs to ref category", logFilePath);
                        if (Program.part_Asm_SM_Weld_FilesToProcess.Contains(itemIDofFile.ToLower()) == false)
                        {
                            Utility.Log("LTC_PROCESS_PSM:- " + itemIDofFile + " belongs to ref category with ID. Skipping", logFilePath);
                            Program.totalFilesProcessed++;
                            continue;
                        }
                        else
                        {
                            Utility.Log("LTC_PROCESS_PSM:- " + itemIDofFile + " does not belong to ref category with ID", logFilePath);
                        }
                    }
                    else
                        Utility.Log("LTC_PROCESS_PSM:- " + partFileWoExtn + " does not belong to ref category.", logFilePath);
                }
                else
                    Utility.Log("\n LTC_PROCESS_PSM:- SEAssyOut.txt is not available ", logFilePath);

            String Folder = outputFolderPath;

                
                    SOL_SE_TRANSLATE PsmTranslate = new SOL_SE_TRANSLATE();
                    PsmTranslate.SavePsmAsSTATThread(psmFileFullPath, "DXF", logFilePath, Folder, itemIDofFile);


                    Utility.Log(psmFile + " Translation is Completed..", logFilePath);
                    Program.totalFilesProcessed++;

                    String ProjectID = LTC_SE_SET_PROPERTIES.getProperty("SHEETMETAL", psmFile, "PROJECT");
                    if (ProjectID == null || ProjectID.Equals("") == true || ProjectID.Equals(" ") == true)
                    {
                        Console.WriteLine("LTC_PROCESS_PSM:- Unable to get the ProjectID/No ProjectID for" + psmFileFullPath);
                        Program.totalFilesProcessed++;
                        continue;
                    }

                    String[] ProjectIdArray = ProjectID.Split(',');
                    String fileName = Path.GetFileNameWithoutExtension(psmFileFullPath);

                    String dxfFilePath = Path.Combine(Folder, fileName + ".dxf");

                    foreach (String projectID in ProjectIdArray)
                    {
                        Console.WriteLine("LTC_PROCESS_PSM:- ProjectID:" + projectID);


                        String dxfFileNewPath = Path.Combine(Folder, fileName + "-" + projectID + ".dxf");
                        if (File.Exists(dxfFilePath) == true)
                        {
                            File.Copy(dxfFilePath, dxfFileNewPath);
                        }
                    }

                    Utility.Log("Deleting the Original DXF file.." + dxfFilePath, logFilePath);
                    try
                    {
                        File.Delete(dxfFilePath);
                    }
                    catch (Exception ex)
                    {
                        Utility.Log("Exception in Deleting the Original DXF file.." + dxfFilePath, logFilePath);
                    }

                
                
            }
        }


        public static void PROCESS_SHEETMETAL_To_STP(String StageDir, String outputFolderPath, String logFilePath)
        {
            string[] PartFiles = Directory.GetFiles(StageDir, "*")
                                             .Select(path => Path.GetFullPath(path))
                                             .Where(x => (x.EndsWith(".psm") || x.EndsWith(".PSM")))
                                             .ToArray();
            if (PartFiles == null || PartFiles.Length == 0)
            {
                Console.WriteLine("LTC_PROCESS_PSM:- NO SHEETMETAL FILES IDENTIFIED");
                return;
            }

            foreach (String psmFileFullPath in PartFiles)
            {
                Utility.Log(psmFileFullPath + " Translation Starting..", logFilePath);

                // 1 AUG 2019 - Dont process ref files
                String psmFile = Path.GetFileName(psmFileFullPath);
                string psmFileName = Path.GetFileName(psmFileFullPath);
                string itemIDofFile = LTC_SE_SET_PROPERTIES.getProperty("SHEETMETAL", psmFileName, "ITEM_ID");
                if (Program.part_Asm_SM_Weld_FilesToProcess.Count != 0)
                {
                    string partFileWoExtn = Path.GetFileNameWithoutExtension(psmFileFullPath);
                    Utility.Log("\n LTC_PROCESS_PSM:- " + partFileWoExtn + "Checking if part belongs to ref category", logFilePath);
                    if (Program.part_Asm_SM_Weld_FilesToProcess.Contains(partFileWoExtn.ToLower()) == false)
                    {
                        Utility.Log("LTC_PROCESS_PSM:- " + partFileWoExtn + " belongs to ref category. Skipping", logFilePath);
                        Utility.Log("LTC_PROCESS_PSM:- " + partFileWoExtn + " Making sure part belongs to ref category with ID", logFilePath);
                        Utility.Log("LTC_PROCESS_PSM:- " + "checking if " + itemIDofFile + " belongs to ref category", logFilePath);
                        if (Program.part_Asm_SM_Weld_FilesToProcess.Contains(itemIDofFile.ToLower()) == false)
                        {
                            Utility.Log("LTC_PROCESS_PSM:- " + itemIDofFile + " belongs to ref category with ID. Skipping", logFilePath);
                            Program.totalFilesProcessed++;
                            continue;
                        }
                        else
                        {
                            Utility.Log("LTC_PROCESS_PSM:- " + itemIDofFile + " does not belong to ref category with ID", logFilePath);
                        }
                    }
                    else
                        Utility.Log("LTC_PROCESS_PSM:- " + partFileWoExtn + " does not belong to ref category.", logFilePath);
                }
                else
                    Utility.Log("\n LTC_PROCESS_PSM:- SEAssyOut.txt is not available ", logFilePath);

                String Folder = outputFolderPath;


                SOL_SE_TRANSLATE PsmTranslate = new SOL_SE_TRANSLATE();
                PsmTranslate.SavePsmAsSTATThread(psmFileFullPath, "STP", logFilePath, Folder, itemIDofFile);


                Utility.Log(psmFile + " Translation is Completed..", logFilePath);
                Program.totalFilesProcessed++;

                



            }
        }
    }
}
