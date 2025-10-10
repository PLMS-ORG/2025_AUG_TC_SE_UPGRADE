using LTC_SE2CACHE;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LTC_SE2CACHE
{
    class SE_PDF_UPLOAD
    {
        public static void Pdf_Upload(String StageDir,String logFilePath)
        {   
            if (collectFilesAndCallIpsUpload(StageDir, logFilePath) == false)
            {
                Utility.Log("SE_PDF_UPLOAD:- IPS UPLOAD OF PG PDF FAILED", logFilePath);
                Console.WriteLine("SE_PDF_UPLOAD:- IPS UPLOAD OF PG PDF FAILED");
                return;
            }            
        }

        public static bool collectFilesAndCallIpsUpload(String folderPath, String logFilePath)
        {
            List<String[]> ipsTextList = new List<string[]>();

            Console.WriteLine("SE_PDF_UPLOAD:- FOLDERPATH SEARCHING FOR PDF: " + folderPath);
            string[] PDFFiles = Directory.GetFiles(folderPath, "*")
                                         .Select(path => Path.GetFileName(path))
                                         .Where(x => (x.EndsWith(".pdf") || x.EndsWith(".PDF")))
                                         .ToArray();
            if (PDFFiles == null || PDFFiles.Length == 0)
            {
                Console.WriteLine("SE_PDF_UPLOAD:- NO PDF FILES IDENTIFIED");
                return false;
            }

            Console.WriteLine("SE_PDF_UPLOAD:- NO OF PDF FILES IDENTIFIED: " +PDFFiles.Length );
            foreach (String pdfFile in PDFFiles)
            {
                String fileNameWOExtn = "";
                String projectid="";
                fileNameWOExtn = Path.GetFileNameWithoutExtension(pdfFile); 
                if (pdfFile.Contains("-") == true)
                {
                    var pdfFileArray = fileNameWOExtn.Split(new[] { '-' }, 2);
                    //String[] pdfFileArray = fileNameWOExtn.Split('-');
                    if (pdfFileArray == null || pdfFileArray.Length == 0) continue;
                    fileNameWOExtn = pdfFileArray[0];
                    projectid=pdfFileArray[1];
                }
                else
                {

                    fileNameWOExtn = Path.GetFileNameWithoutExtension(pdfFile);
                    if (fileNameWOExtn == null || fileNameWOExtn.Equals("") == true)
                    {
                        Console.WriteLine("SE_PDF_UPLOAD:- Unable to get the Item ID and Revision for " + pdfFile);
                        Program.totalDeliverablesUploaded++;
                        continue;
                    }
                }

                //String[] fileNameWOExtnArr = fileNameWOExtn.Split('_');                
                String itemID = "";
                String RevID = "";


                
                String correspondingDftFile = fileNameWOExtn + ".dft";
                itemID = LTC_SE_SET_PROPERTIES.getProperty("DRAFT", correspondingDftFile, "ITEM_ID");
                if (itemID == null || itemID.Equals("") == true)
                {
                    Console.WriteLine("SE_PDF_UPLOAD:- Unable to get the Item ID for" + fileNameWOExtn + ".dft");
                    Program.totalDeliverablesUploaded++;
                    continue;
                }

                RevID = LTC_SE_SET_PROPERTIES.getProperty("DRAFT", correspondingDftFile, "REV_ID");
                if (RevID == null || RevID.Equals("") == true)
                {
                    Console.WriteLine("SE_PDF_UPLOAD:- Unable to get the RevID for" + fileNameWOExtn + ".dft");
                    Program.totalDeliverablesUploaded++;
                    continue;
                }

                /* String item_id, String rev, String ItemName,String dsName, String fileRef,String file, String datasetDescription, String relationName*/
                Console.WriteLine("ItemID: " + itemID + " RevID: " + RevID);
                string[] ipsLine = new string[9];
                String itemName = itemID + "/" + RevID;
                String dsName = "";
                String fileRef = "PDF_Reference";
                String pdfFile1 = "";
                if (pdfFile.Contains("-") == true)
                {
                    pdfFile1 = itemID + RevID + "-" + projectid + ".pdf";
                    dsName = itemID + "/" + RevID + "-" + projectid;
                }
                else
                {
                    pdfFile1 = itemID + RevID + ".pdf";
                    dsName = itemID + "/" + RevID;
                }
                String PDFFilePath = Path.Combine(folderPath, pdfFile);
                String PDFFilePath1 = Path.Combine(folderPath, pdfFile1);
                Console.WriteLine("PDFFilePath: " + PDFFilePath);
                Console.WriteLine("PDFFilePath1: " + PDFFilePath1);
                // 4 SEPT -- ITEMIDREVID.pdf already exists in the directory.
                // Hence no need to Move(Rename) the FILE.
                if (PDFFilePath.Equals(PDFFilePath1, StringComparison.OrdinalIgnoreCase) == false)
                {
                    if (File.Exists(PDFFilePath1))
                    {
                        try
                        {
                            File.Delete(PDFFilePath1);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("SE_PDF_UPLOAD:- DELETE OF: " + PDFFilePath1 + " FAILED");
                            Program.totalDeliverablesUploaded++;
                            continue;
                        }
                    }
                    try
                    {
                        File.Move(PDFFilePath, PDFFilePath1);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("SE_PDF_UPLOAD:- MOVE FAILED FOR,SKIPPING PDF_UPLOAD FOR: " + ex.Message + PDFFilePath);
                        Program.totalDeliverablesUploaded++;
                        continue;
                    }
                }
                else
                {
                    Console.WriteLine("SE_PDF_UPLOAD:- NO NEED TO RENAME THE PDF FILE. ALREADY FOLLOWS NAMING RULE " + PDFFilePath);
                }
                
                String dsDesc = "LTC" + "_PDF";
                String relationName = "";
                relationName = "IMAN_manifestation";
                
                ipsLine[0] = itemID;
                ipsLine[1] = RevID;
                ipsLine[2] = itemName;
                ipsLine[3] = dsName;
                ipsLine[4] = fileRef;
                ipsLine[5] = PDFFilePath1;
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

            SE_IPS_UPLOAD.CreateIPSUploadTEXTFile(folderPath, "PDF", ipsTextList, logFilePath);
            return true;
            
        }

    }
}
