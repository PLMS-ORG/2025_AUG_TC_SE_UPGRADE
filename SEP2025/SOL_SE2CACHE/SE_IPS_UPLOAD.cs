using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LTC_SE2CACHE
{
    class SE_IPS_UPLOAD
    {

        //public static int CreateIPSUploadTEXTFile(String taskResultDir, String item_id, String rev, String ItemName,
        //String dsName, String fileRef,
        //String file, String datasetDescription, String relationName, String logFilePath)
        //{
        //    Utility.Log("IPSUpload: creating text file", logFilePath);

        //    String firstLine = "";
        //    String SecondLine = "";

        //    String FULLipsFileName = System.IO.Path.Combine(taskResultDir, "ips_" + item_id + "_" + rev + ".txt");
        //    System.IO.StreamWriter fileStream =
        //        new System.IO.StreamWriter(FULLipsFileName);


        //    firstLine = "!~ItemID~RevID~Name~DsetName~FileRef~File~DsetDesc~RelationName";
        //    Console.WriteLine(firstLine);
        //    SecondLine = item_id + "~" + rev + "~" + ItemName + "~" + dsName + "~" + fileRef + "~" + file + "~" + datasetDescription + "~" + relationName;

        //    Console.WriteLine(SecondLine);

        //    fileStream.WriteLine(firstLine);
        //    fileStream.WriteLine(SecondLine);
        //    fileStream.Close();

        //    RunIPSUpload(logFilePath,"", FULLipsFileName);
        //    return 0;
        //}

        public static int CreateIPSUploadTEXTFile(String taskResultDir, String DsType, List<String[]> secondLine, String logFilePath)
        {
            Utility.Log("IPSUpload: creating text file", logFilePath);

            if (secondLine.Count == 0)
            {
                Utility.Log("IPSUpload: No IPSUpload TEXT FOUND", logFilePath);
                return -1;
            }

            //String firstLine = "";
            String SecondLine = "";

            String FULLipsFileName = System.IO.Path.Combine(taskResultDir, "ips_" + "SOL_" + DsType + ".txt");
            System.IO.StreamWriter fileStream =
                new System.IO.StreamWriter(FULLipsFileName);


            //firstLine = "!~ItemID~RevID~Name~DsetName~FileRef~File~DsetDesc~RelationName";
            //Console.WriteLine("IPSUpload:- " + firstLine);
            //fileStream.WriteLine(firstLine);
            String dsRealName = "";
            if (DsType.Equals("PDF", StringComparison.OrdinalIgnoreCase) == true)
            {
                dsRealName = "PDF";
            }
            else if (DsType.Equals("DXF", StringComparison.OrdinalIgnoreCase) == true)
            {
                dsRealName = "DXF";
            }
            else if (DsType.Equals("STP", StringComparison.OrdinalIgnoreCase) == true)
            {
                dsRealName = "AI4_STEP";
            }

            foreach (String[] sLine in secondLine)
            {
                SecondLine = sLine[0] + "~" + sLine[1] + "~" + sLine[2] + "~" + sLine[3] + "~" + sLine[4] + "~" + sLine[5] + "~" + sLine[6] + "~" + sLine[7]
                    + "~" + sLine[8];
                //SecondLine = item_id + "~" + rev + "~" + ItemName + "~" + dsName + "~" + fileRef + "~" + file + "~" + datasetDescription + "~" + relationName;
                Console.WriteLine("3DEUpload:- " + SecondLine);
                Run3DEUpload(sLine[0], sLine[1], sLine[3], dsRealName, sLine[7], sLine[4], sLine[5], sLine[0] + sLine[1] + "." + DsType.ToLower(), sLine[8], logFilePath);
                //fileStream.WriteLine(SecondLine);
            }

            //fileStream.Close();

            //RunIPSUpload(logFilePath, DsType, FULLipsFileName);
            return 0;
        }

        public static void RunIPSUpload(String logFilePath, String DsType, String FULLIPSUPLOADTEXTFILE)
        {
            Utility.Log("RunIPSUpload- IPSUPLOAD.CMD ", logFilePath);
            if (DsType.Equals("") == true || DsType == null)
            {
                Console.WriteLine("RunIPSUpload - DataSetType is Empty");
                return;
            }

            String IPSUPLOADPATH = "";
            try
            {
                IPSUPLOADPATH = System.Environment.GetEnvironmentVariable("MODULE_HOME");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to get IPS_UPLOAD_CMD_HOME Path");
                Console.WriteLine(ex.Message);
                return;
            }

            Console.WriteLine("MODULE_HOME" + IPSUPLOADPATH);
            String IPSUPLOADCMDFILE = "";
            if (DsType.Equals("PDF", StringComparison.OrdinalIgnoreCase) == true)
            {
                IPSUPLOADCMDFILE = "IPSUPLOAD.CMD";
            }
            else if (DsType.Equals("DXF", StringComparison.OrdinalIgnoreCase) == true)
            {
                IPSUPLOADCMDFILE = "IPSUPLOAD_DXF.CMD";
            }
            else if (DsType.Equals("STP", StringComparison.OrdinalIgnoreCase) == true)
            {
                IPSUPLOADCMDFILE = "IPSUPLOAD_STP.CMD";
            }

            // Use ProcessStartInfo class
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = true;
            startInfo.FileName = IPSUPLOADCMDFILE;
            //startInfo.FileName = Path.Combine(solidEdgetranslationservicePath, solidEdgeEXE);
            startInfo.WindowStyle = ProcessWindowStyle.Normal;
            startInfo.Arguments = "\"" + FULLIPSUPLOADTEXTFILE + "\"";
            startInfo.WorkingDirectory = IPSUPLOADPATH;


            if (File.Exists(Path.Combine(IPSUPLOADPATH, IPSUPLOADCMDFILE)) == false)
            {
                Utility.Log("IPSUPLOAD EXE not Found", logFilePath);
                return;
            }

            try
            {
                // Start the process with the info we specified.
                // Call WaitForExit and then the using statement will close.
                using (Process exeProcess = Process.Start(startInfo))
                {
                    exeProcess.WaitForExit();
                    //exeProcess.StandardOutput.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                Utility.Log("Exception in Process.Start", logFilePath);
                Utility.Log(ex.Message, logFilePath);
            }

        }

        //3DE_Upload_Dataset.exe %upg% -i=%ItemId% -r=%revid% -dsName=%ItemId%/%revid% -dsType=PDF -dsRelation=IMAN_specification -fileRef=PDF_Reference -filePath=%stagedir%\%ItemId%%revid%.pdf -log=%stagedir%\3deupload_%ItemId%_%revid%.log -dsNamedRefName=%ItemId%%revid%.pdf
        public static void Run3DEUpload(String itemId, String revId, String dsName, String dsType, String dsRelation, String fileRef, String filePath, String dsNamedRefName, String projectCode, String logFilePath)
        {
            String UTILUPLOADPATH = "";

            UTILUPLOADPATH = Utility.GetClickOnceLocation();

            if (UTILUPLOADPATH==null || UTILUPLOADPATH.Length == 0)
            {
                Console.WriteLine("Unable to get 3DE_UPLOAD_CMD_HOME Path");
                return;
            }

            Console.WriteLine("MODULE_HOME" + UTILUPLOADPATH);

            // Use ProcessStartInfo class
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = true;
            startInfo.FileName = "3DE_Upload_Dataset_TC2506.exe";
            //startInfo.FileName = Path.Combine(solidEdgetranslationservicePath, solidEdgeEXE);
            startInfo.WindowStyle = ProcessWindowStyle.Normal;
            //String args = "-u=dcproxy -p=dcproxy -g=dba" + " -i=" + itemId + " -r=" + revId + " -dsType=" + dsType + " -dsName=" + dsName + " -dsRelation=" + dsRelation + " -fileRef=" + fileRef + " -filePath=" + filePath + " -log=" + logFilePath + " -dsNamedRefName=" + dsNamedRefName
            //    + " -projectCode=" + projectCode;

            // 29-04-2025 | Murali | No need to assign Project ID during the 3DE Upload | Request from Sanju and Kris - LTC
            String args = "-u=dcproxy -p=dcproxy -g=dba" + " -i=" + itemId + " -r=" + revId + " -dsType=" + dsType + " -dsName=" + dsName + " -dsRelation=" + dsRelation + " -fileRef=" + fileRef + " -filePath=" + filePath + " -log=" + logFilePath + " -dsNamedRefName=" + dsNamedRefName;
            Console.WriteLine("Args: " + args);
            startInfo.Arguments = args;
            startInfo.WorkingDirectory = UTILUPLOADPATH;

            if (File.Exists(Path.Combine(UTILUPLOADPATH, "3DE_Upload_Dataset_TC2506.exe")) == false)
            {
                Utility.Log("3DEUPLOAD EXE not Found", logFilePath);
                return;
            }

            try
            {
                // Start the process with the info we specified.
                // Call WaitForExit and then the using statement will close.
                using (Process exeProcess = Process.Start(startInfo))
                {
                    exeProcess.WaitForExit();
                    //exeProcess.StandardOutput.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                Utility.Log("Exception in Process.Start", logFilePath);
                Utility.Log(ex.Message, logFilePath);
            }

        }

    }
}
