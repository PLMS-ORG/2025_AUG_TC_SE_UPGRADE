using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LTC_SE2CACHE
{
    class SOL_STP_UPLOAD
    {
        public static void Stp_Upload(String stpOutputFolder,String stageDir, String logFilePath)
        {


            if (collectFilesAndCallIpsUpload(stpOutputFolder, stageDir,logFilePath) == false)
            {
                Utility.Log("SOL_STP_UPLOAD:- IPS UPLOAD OF STP FAILED", logFilePath);
                Console.WriteLine("SOL_STP_UPLOAD:- IPS UPLOAD OF STP FAILED");
                return;
            }
        }

        public static bool collectFilesAndCallIpsUpload(String folderPath, String stageDir, String logFilePath)
        {
            List<String[]> ipsTextList = new List<string[]>();

            string[] STPFiles = Directory.GetFiles(folderPath, "*")
                                         .Select(path => Path.GetFileName(path))
                                         .Where(x => (x.EndsWith(".stp") || x.EndsWith(".STP")))
                                         .ToArray();
            if (STPFiles == null || STPFiles.Length == 0)
            {
                Console.WriteLine("SOL_STP_UPLOAD:- NO STP FILES IDENTIFIED");
                return false;
            }

            Console.WriteLine("SOL_STP_UPLOAD:- NO OF STP FILES IDENTIFIED: " + STPFiles.Length);
            foreach (String STPFile in STPFiles)
            {
                String fileNameWOExtn = Path.GetFileNameWithoutExtension(STPFile);
                if (fileNameWOExtn == null || fileNameWOExtn.Equals("") == true)
                {
                    Console.WriteLine("SOL_STP_UPLOAD:- Unable to get the Item ID and Revision for " + STPFile);
                    Program.totalDeliverablesUploaded++;
                    continue;
                }

                //String[] fileNameWOExtnArr = fileNameWOExtn.Split('_');                
                String itemID = "";
                String RevID = "";


                //if (fileNameWOExtnArr == null || fileNameWOExtnArr.Length == 0)
                //{
                //    Console.WriteLine("SOL_PROCESS_DRAFT:- Unable to get the Item ID and Revision for " + STPFile);
                //    continue;

                //}

                //if (fileNameWOExtnArr.Length == 2)
                //{
                //    //DrawingPrefix = fileNameWOExtnArr[0];
                //    itemID = fileNameWOExtnArr[0];
                //    RevID = fileNameWOExtnArr[1];
                //    //RevID = RevID.Substring(1);
                //}
                //else if (fileNameWOExtnArr.Length == 3)
                //{
                //    //DrawingPrefix = fileNameWOExtnArr[0];
                //    itemID = fileNameWOExtnArr[0] + "_" + fileNameWOExtnArr[1];
                //    RevID = fileNameWOExtnArr[2];
                //}
                //else
                //{
                //    Console.WriteLine("SOL_PROCESS_DRAFT:- Unable to Split FileName to GET ItemID and RevID " + STPFile);
                //    continue;
                //}

              

                
                String extension = Utility.getMatchingFile(STPFile, ".stp", stageDir);
                if (extension == null || extension.Equals("") == true)
                {
                    Console.WriteLine("SE_STP_UPLOAD:- Unable to get the Extension for" + fileNameWOExtn + ".stp");
                    Program.totalDeliverablesUploaded++;
                    continue;

                }

                String correspondingFile = fileNameWOExtn;
                String dsType = "";
                if (extension.Equals(".asm", StringComparison.OrdinalIgnoreCase) == true)
                {
                    dsType = "ASSEMBLY";
                    correspondingFile = correspondingFile + ".asm";

                }
                else if (extension.Equals(".par", StringComparison.OrdinalIgnoreCase) == true)
                {
                    dsType = "PART";
                    correspondingFile = correspondingFile + ".par";
                }
                else if (extension.Equals(".pwd", StringComparison.OrdinalIgnoreCase) == true)
                {
                    dsType = "WELDMENT";
                    correspondingFile = correspondingFile + ".pwd";
                }
                else if (extension.Equals(".psm", StringComparison.OrdinalIgnoreCase) == true)
                {
                    dsType = "SHEETMETAL";
                    correspondingFile = correspondingFile + ".psm";

                }
                RevID = LTC_SE_SET_PROPERTIES.getProperty(dsType, correspondingFile, "REV_ID");
                if (RevID == null || RevID.Equals("") == true)
                {
                    Console.WriteLine("SE_STP_UPLOAD:- Unable to get the RevID for" + correspondingFile);
                    Program.totalDeliverablesUploaded++;
                    continue;
                }

                itemID = LTC_SE_SET_PROPERTIES.getProperty(dsType, correspondingFile, "ITEM_ID");
                if (itemID == null || itemID.Equals("") == true)
                {
                    Console.WriteLine("SE_STP_UPLOAD:- Unable to get the itemID for" + correspondingFile);
                    Program.totalDeliverablesUploaded++;
                    continue;
                }
                /* String item_id, String rev, String ItemName,String dsName, String fileRef,String file, String datasetDescription, String relationName*/

                
                string[] ipsLine = new string[9];
                String itemName = itemID + "/" + RevID;
                String dsName = itemID + "/" + RevID;
                String fileRef = "AI4_STP_Reference";
                String STPFile1 = itemID + RevID + ".stp"; // modified for SOLERAS Requirement
                String STPFilePath = Path.Combine(folderPath, STPFile);
                String STPFilePath1 = Path.Combine(folderPath, STPFile1);
                // 4 SEPT -- ITEMIDREVID.stp already exists in the directory.
                // Hence no need to Move(Rename) the FILE.
                if (STPFilePath.Equals(STPFilePath1,StringComparison.OrdinalIgnoreCase) == false)
                {
                    if (File.Exists(STPFilePath1))
                    {
                        try
                        {
                            File.Delete(STPFilePath1);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("SE_STP_UPLOAD:- DELETE OF: " + STPFilePath1 + " FAILED");
                            Program.totalDeliverablesUploaded++;
                            continue;
                        }
                    }
                    try
                    {
                        File.Move(STPFilePath, STPFilePath1);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("SE_STP_UPLOAD:- MOVE FAILED FOR,SKIPPING PDF_UPLOAD FOR: " + STPFilePath1);
                        Program.totalDeliverablesUploaded++;
                        continue;
                    }
                }
                else
                {
                    Console.WriteLine("SE_STP_UPLOAD:- NO NEED TO RENAME THE STP FILE. ALREADY FOLLOWS NAMING RULE " + STPFilePath1);
                }
               
                String dsDesc = "SOL" + "_STP";
                String relationName = "";                
                relationName = "IMAN_specification";
                
                ipsLine[0] = itemID;
                ipsLine[1] = RevID;
                ipsLine[2] = itemName;
                ipsLine[3] = dsName;
                ipsLine[4] = fileRef;
                ipsLine[5] = STPFilePath1;
                ipsLine[6] = dsDesc;
                ipsLine[7] = relationName;
                ipsLine[8] = "NA";
                ipsTextList.Add(ipsLine);

                Program.totalDeliverablesUploaded++;

            }

            SE_IPS_UPLOAD.CreateIPSUploadTEXTFile(folderPath, "STP", ipsTextList, logFilePath);
            return true;

        }

    }
}
