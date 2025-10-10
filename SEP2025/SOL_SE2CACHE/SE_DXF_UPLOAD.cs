using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LTC_SE2CACHE
{
    class SE_DXF_UPLOAD
    {

        public static void Dxf_Upload(String StageDir, String logFilePath)
        {   
            if (collectFilesAndCallIpsUpload(StageDir,  logFilePath) == false)
            {
                Utility.Log("SE_DXF_UPLOAD:- IPS UPLOAD OF DCH DXF FAILED", logFilePath);
                Console.WriteLine("SE_DXF_UPLOAD:- IPS UPLOAD OF DCH DXF FAILED");
                return;
            }
        }

        public static bool collectFilesAndCallIpsUpload(String folderPath, String logFilePath)
        {
            List<String[]> ipsTextList = new List<string[]>();

            string[] DXFFiles = Directory.GetFiles(folderPath, "*")
                                         .Select(path => Path.GetFileName(path))
                                         .Where(x => (x.EndsWith(".dxf") || x.EndsWith(".DXF")))
                                         .ToArray();
            if (DXFFiles == null || DXFFiles.Length == 0)
            {
                Console.WriteLine("SE_DXF_UPLOAD:- NO DXF FILES IDENTIFIED");
                return false;
            }

            Console.WriteLine("SE_DXF_UPLOAD:- NO OF DXF FILES IDENTIFIED: " + DXFFiles.Length);
            foreach (String dxfFile in DXFFiles)
            {
                String projectid = "";
                String fileNameWOExtn = Path.GetFileNameWithoutExtension(dxfFile);
                if (fileNameWOExtn == null || fileNameWOExtn.Equals("") == true)
                {
                    Console.WriteLine("SE_DXF_UPLOAD:- Unable to get the Item ID and Revision for " + dxfFile);
                    Program.totalDeliverablesUploaded++;
                    continue;
                }

                if (dxfFile.Contains("-") == true)
                {

                    String[] dxfFileArray = fileNameWOExtn.Split('-');
                    if (dxfFileArray == null || dxfFileArray.Length == 0) continue;
                    fileNameWOExtn = dxfFileArray[0];
                    projectid = dxfFileArray[1];
                }
                else
                {

                    fileNameWOExtn = Path.GetFileNameWithoutExtension(dxfFile);
                    if (fileNameWOExtn == null || fileNameWOExtn.Equals("") == true)
                    {
                        Console.WriteLine("SE_DXF_UPLOAD:- Unable to get the Item ID and Revision for " + dxfFile);
                        Program.totalDeliverablesUploaded++;
                        continue;
                    }
                }

                //String[] fileNameWOExtnArr = fileNameWOExtn.Split('_');                
                String itemID = "";
                String RevID = "";


                //if (fileNameWOExtnArr == null || fileNameWOExtnArr.Length == 0)
                //{
                //    Console.WriteLine("SOL_PROCESS_DRAFT:- Unable to get the Item ID and Revision for " + dxfFile);
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
                //    Console.WriteLine("SOL_PROCESS_DRAFT:- Unable to Split FileName to GET ItemID and RevID " + dxfFile);
                //    continue;
                //}

                String correspondingDftFile = fileNameWOExtn + ".dft";
                itemID = LTC_SE_SET_PROPERTIES.getProperty("DRAFT", correspondingDftFile, "ITEM_ID");
                if (itemID == null || itemID.Equals("") == true)
                {
                    Console.WriteLine("SE_DXF_UPLOAD:- Unable to get the Item ID for" + fileNameWOExtn + ".dft");
                    Program.totalDeliverablesUploaded++;
                    continue;
                }

                RevID = LTC_SE_SET_PROPERTIES.getProperty("DRAFT", correspondingDftFile, "REV_ID");
                if (RevID == null || RevID.Equals("") == true)
                {
                    Console.WriteLine("SE_DXF_UPLOAD:- Unable to get the RevID for" + fileNameWOExtn + ".dft");
                    Program.totalDeliverablesUploaded++;
                    continue;
                }

                /* String item_id, String rev, String ItemName,String dsName, String fileRef,String file, String datasetDescription, String relationName*/

                string[] ipsLine = new string[9];
                String itemName = itemID + "/" + RevID;
                String dsName = itemID + "/" + RevID;
                String fileRef = "DXF";
                String DxfFile1 = itemID + RevID + ".dxf";

                if (dxfFile.Contains("-") == true)
                {
                    
                   // DxfFile1 = itemID + RevID + "_" + projectid + ".dxf";
                   //dsName = itemID + "/" + RevID + "-" + projectid;

                    // Sindhuja | 23-05-2025 | ProjectID is not needed in DXF file Name
                    DxfFile1 = itemID + RevID +".dxf";
                    dsName = itemID + "/" + RevID ;
                }
                else
                {
                    DxfFile1 = itemID + RevID + ".dxf";
                    dsName = itemID + "/" + RevID;
                }

                String DXFFilePath = Path.Combine(folderPath, dxfFile);
                String DXFFilePath1 = Path.Combine(folderPath, DxfFile1);

                // 4 SEPT -- ITEMIDREVID.dxf already exists in the directory.
                // Hence no need to Move(Rename) the FILE.
                if (DXFFilePath.Equals(DXFFilePath1, StringComparison.OrdinalIgnoreCase) == false)
                {
                    if (File.Exists(DXFFilePath1))
                    {
                        try
                        {
                            File.Delete(DXFFilePath1);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("SE_DXF_UPLOAD:- DELETE OF: " + DxfFile1 + " FAILED");
                            Program.totalDeliverablesUploaded++;
                            continue;
                        }
                    }
                    try
                    {
                        File.Move(DXFFilePath, DXFFilePath1);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("SE_DXF_UPLOAD:- MOVE FAILED FOR,SKIPPING PDF_UPLOAD FOR: " + DXFFilePath);
                        Program.totalDeliverablesUploaded++;
                        continue;
                    }
                }
                else
                {
                    Console.WriteLine("SE_DXF_UPLOAD:- NO NEED TO RENAME THE DXF FILE. ALREADY FOLLOWS NAMING RULE " + DXFFilePath);
                }

                String dsDesc = "SOL" + "_DXF";
                String relationName = "";                                
                relationName = "IMAN_manifestation";
                
                ipsLine[0] = itemID;
                ipsLine[1] = RevID;
                ipsLine[2] = itemName;
                ipsLine[3] = dsName;
                ipsLine[4] = fileRef;
                ipsLine[5] = DXFFilePath1;
                ipsLine[6] = dsDesc;
                ipsLine[7] = relationName;
                if (projectid == null || projectid.Equals("") == true)
                {
                    projectid = "";
                }
                ipsLine[8] = projectid;
                ipsTextList.Add(ipsLine);

                Program.totalDeliverablesUploaded++;

            }

            SE_IPS_UPLOAD.CreateIPSUploadTEXTFile(folderPath, "DXF", ipsTextList, logFilePath);
            return true;

        }

    }
}
