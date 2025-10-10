using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LTC_SE2CACHE
{
    // 26 -  SEPT
    class SOL_WaterMarkPDF
    {

        public static void RunWaterMarkPDFUtility(String logFilePath, String stageDirectory, String pdfFilePath)
        {
            
            pdfFilePath = Path.Combine(stageDirectory, pdfFilePath);
            String ReleasedPNGPath = Path.Combine(stageDirectory, "Released.png");

            String WATERMARKPDFCMDFILE = "PDFWaterMark.bat";

            Utility.Log("RunWaterMarkPDFUtility- PDFWaterMark.bat ", logFilePath);
            if (stageDirectory.Equals("") == true || stageDirectory == null)
            {
                Console.WriteLine("RunWaterMarkPDFUtility - stageDirectory is Empty");
                return;
            }

            String ModuleHome = "";
            try
            {
                ModuleHome = System.Environment.GetEnvironmentVariable("MODULE_HOME");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to get IPS_UPLOAD_CMD_HOME Path");
                Console.WriteLine(ex.Message);
                return;
            }

            Console.WriteLine("MODULE_HOME" + ModuleHome);
            Utility.Log("RunWaterMarkPDFUtility- " + pdfFilePath, logFilePath);
            Utility.Log("RunWaterMarkPDFUtility- " + WATERMARKPDFCMDFILE, logFilePath);
            Utility.Log("RunWaterMarkPDFUtility- " + ReleasedPNGPath, logFilePath);

            // Use ProcessStartInfo class
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.CreateNoWindow = false;
            startInfo.UseShellExecute = true;
            startInfo.FileName = WATERMARKPDFCMDFILE;
            //startInfo.FileName = Path.Combine(solidEdgetranslationservicePath, solidEdgeEXE);
            startInfo.WindowStyle = ProcessWindowStyle.Normal;
            startInfo.Arguments = "\"" + pdfFilePath + "\"" + " " + "\"" + ReleasedPNGPath + "\"";
            startInfo.WorkingDirectory = ModuleHome;


            if (File.Exists(Path.Combine(ModuleHome, WATERMARKPDFCMDFILE)) == false)
            {
                Utility.Log("WATERMARKPDFCMDFILE BAT not Found", logFilePath);
                return;
            }

            try
            {
                // Start the process with the info we specified.
                // Call WaitForExit and then the using statement will close.
                using (Process exeProcess = Process.Start(startInfo))
                {
                    Utility.Log("RunWaterMarkPDFUtility: " + "Running Process: ", logFilePath);
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
