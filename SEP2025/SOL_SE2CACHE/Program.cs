using SolidEdge.SDK;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Threading;

namespace LTC_SE2CACHE
{
    /**
     * Murali - 03-NOV-2020 Un-managed SE2CACHE
     * Murali - 03-MAR-2020 Enable functionality to save dxf from flat pattern features in psm files.
     **/
    class Program
    {
        static String LICENSENOTIFICATION = "UserLicenseAvailabilityNotification" + ".txt";
        public static List<string> draftFilesToProcess = new List<string>(), part_Asm_SM_Weld_FilesToProcess = new List<string>();
        public static string itemIDRecieved = null, revIDReceived = null;
        public static StreamWriter sw = null;
        public static string failureFilePath = "";

        [STAThread]
        static void Main(string[] args)
        {
            /**
             * TaskFolderPath - Mandatory - Folder where SE files are downloaded.
             * output Dir - Mandatory - Folder where PDF, DXF , STP files are generated.
             * Network Folder to Copy PDF - Not Mandatory - PDF files can be copied to a network share Folder.
             * LogFile Directory - Mandatory - Folder where the Log Files need to be generated
             * IgnoreReleaseStatus - Mandatory - Yes/No - Ignores Files with Release Status - yes.
             * CreatePDF - Mandatory - Yes/No -createPDF=Yes
             * CreateSTP - Mandatory - Yes/No
             * CreateDXF - Mandatory -Yes/No
             * IgnoreRefFiles - Mandatory
             * Stamp-
             * itemID-
             * revID-
             * Network Folder to Copy DXF - Not Mandatory - DXF files can be copied to a network share Folder
             **/
            Console.WriteLine("29-04-2025 | Murali | Do not Rename the PDF using the Project ID | Request from Sanju + Kris at LTC");
            Console.WriteLine("23-05-2025 | Sindhuja | DXF generated for PSM, PDF generated for Draft | Request from Sanju");
            Console.WriteLine("20-06-2025 | Sindhuja | Create STP for PSM | Request from Sanju");
            if (args.Length < 8) 
            {
              //  Console.WriteLine("[TaskFolderPath] [output Dir] [Network Folder to Copy PDF Files] [LogFile Directory] [IgnoreReleaseStatus(Yes/No)] [-createPDF(Yes/No)] [-createDXF(Yes/No)] [-createSTP(Yes/No)] [-ignoreRefFiles(Optional)] [-stamp=StampingText(Optional)] [itemId-optional] [revid-optional] [Network Folder to Copy DXF Files]");
              // Sindhuja||23-06-2024||Remove IgnoreReleaseStatus arugument
                Console.WriteLine("[TaskFolderPath] [output Dir] [Network Folder to Copy PDF Files] [LogFile Directory]  [-createPDF(Yes/No)] [-createDXF(Yes/No)] [-createSTP(Yes/No)] [-ignoreRefFiles(Optional)] [-stamp=StampingText(Optional)] [itemId-optional] [revid-optional] [Network Folder to Copy DXF Files]");
                return;

            }

            String logFilePath = "";
            String stageDir = "";
            string progressLogPath = "";
            // Place the Log file in Stage Directory.
            if (args[3] != null)
            {
                stageDir = args[3];
                logFilePath = System.IO.Path.Combine(stageDir, "LTC_SE2CACHE" + "_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm", CultureInfo.InvariantCulture) + ".txt");
                failureFilePath = System.IO.Path.Combine(stageDir, "failure.txt");
                progressLogPath = System.IO.Path.Combine(stageDir, "ProgressLog.txt");
                sw = new StreamWriter(progressLogPath);                
                Console.WriteLine("Opening LogFile @ {0} " + logFilePath);
                Utility.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);
            }
            else
            {
                logFilePath = "LTC_SE2CACHE" + "_" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm", CultureInfo.InvariantCulture) + ".txt";
                Console.WriteLine("Opening LogFile @ {0} " + logFilePath);
                Utility.Log("Utility Started @ " + System.DateTime.Now.ToString(), logFilePath);
            }

            Utility.Log("---Inputs----", logFilePath);
            Utility.Log("----------------------------", logFilePath);
            string TaskFolderPath = args[0];
            Utility.Log("TaskFolderPath: " + TaskFolderPath, logFilePath);
            string outPutFolderPath = args[1];
            Utility.Log("outPutFolderPath: " + outPutFolderPath, logFilePath);
            string networkFolderPath = args[2];
            Utility.Log("networkFolderPath To Copy PDF Files: " + networkFolderPath, logFilePath);
            string networkFolderPathForDXFCopy = "";
            if (args.Length == 11)
            {
                networkFolderPathForDXFCopy = args[10];
            }
            if (args.Length == 13)
            {
                networkFolderPathForDXFCopy = args[12];
            }

            Utility.Log("networkFolderPath To Copy DXF Files: " + networkFolderPathForDXFCopy, logFilePath);
            Utility.Log("stageDir: " + stageDir, logFilePath);
            string IgnoreReleaseStatus = args[4];
            Utility.Log("IgnoreReleaseStatus: " + IgnoreReleaseStatus, logFilePath);
            //Utility.IgnoreReleaseStatus = IgnoreReleaseStatus;
            Utility.Log("----------------------------", logFilePath);
            if (args.Length == 13)
            {
                itemIDRecieved = args[10];
                revIDReceived = args[11];
            }
            if (itemIDRecieved != null & revIDReceived != null)
                Utility.Log("Item id and rev id to process is received: " + itemIDRecieved + " " + revIDReceived, logFilePath);
            else
                Utility.Log("Item id and rev id to process is not received. " + itemIDRecieved + " " + revIDReceived, logFilePath);


            if (System.IO.Directory.Exists(networkFolderPath) == false || networkFolderPath==null ||
                networkFolderPath.Equals("")==true)
            {
                Console.WriteLine("networkFolderPath does not Exist: " + networkFolderPath);
                return;
            }



            if (stageDir == null)
            {
                Console.WriteLine("LTC_SE2CACHE:- Task Folder is Empty");

                return;
            }

            // Check for LicenseNotification.txt file. If it Exists, Throw an Exception and Exit. As a Result, Dispatcher Should get TERMINAL State.
            String LicenseNotificationFile = System.IO.Path.Combine(TaskFolderPath, LICENSENOTIFICATION);
            if (System.IO.File.Exists(LicenseNotificationFile) == true)
            {
                Utility.Log("LTC_SE2CACHE: EXCEPTION, NO LICENSE TO RUN THIS UTILITY", logFilePath);
                return;
            }


            //Check for ref file
            if (args.Length > 8)
            {
                bool ignoreRefFiles = false;
                for (int i = 8; i < args.Length; i++)
                {
                    if (args[i].Equals("-ignoreRefFiles"))
                        ignoreRefFiles = true;
                }

                if (ignoreRefFiles == true)
                {
                    Utility.Log("LTC_SE2CACHE:      Ref file argument found", logFilePath);
                    string filePath = Path.Combine(stageDir, "SEAssyOut.txt");
                    if (File.Exists(filePath))
                    {
                        string[] fileContents = File.ReadAllLines(filePath);
                        foreach (string s in fileContents)
                        {
                            if (s.Trim().Equals("") == false)
                            {
                                string[] idRev = s.ToLower().Split('/');
                                if (idRev.Length == 2)
                                {
                                    draftFilesToProcess.Add(idRev[0] + idRev[1]);
                                    part_Asm_SM_Weld_FilesToProcess.Add(idRev[0]);
                                }
                                else
                                    Utility.Log("LTC_SE2CACHE: Ref file contains wrong input" + s, logFilePath);
                            }
                        }
                    }
                    else
                        Utility.Log("LTC_SE2CACHE: Ref file SEAssyOut.txt not found", logFilePath);

                    Utility.Log("LTC_SE2CACHE: draftFilesToProcess", logFilePath);
                    foreach (string s in draftFilesToProcess)
                        Utility.Log("LTC_SE2CACHE: " + s, logFilePath);
                    Utility.Log("LTC_SE2CACHE: part_Asm_SM_Weld_FilesToProcess", logFilePath);
                    foreach (string s in part_Asm_SM_Weld_FilesToProcess)
                        Utility.Log("LTC_SE2CACHE: " + s, logFilePath);
                }
            }


            Console.WriteLine("LTC_SE2CACHE:- readCustomProperty" + stageDir);
            LTC_SE_SET_PROPERTIES.readCustomProperty(stageDir);



            if (System.IO.Directory.Exists(outPutFolderPath) == false)
            {
                System.IO.Directory.CreateDirectory(outPutFolderPath);
            }

            bool createPDF = false, createDXF = false, createSTP = false;
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i].Equals("-createPDF=Yes"))
                    createPDF = true;
                else if (args[i].Equals("-createDXF=Yes"))
                    createDXF = true;
                else if (args[i].Equals("-createSTP=Yes"))
                    createSTP = true;
            }

            if (createSTP == false)
            {
                Utility.Log("CreateSTP was not recieved. Considering as PDFDXF_postreleaseIR " + itemIDRecieved + " " + revIDReceived, logFilePath);
                if (args.Length == 12)
                {
                    itemIDRecieved = args[10];
                    revIDReceived = args[11];
                }
                if (itemIDRecieved != null & revIDReceived != null)
                    Utility.Log("Item id and rev id to process is received: " + itemIDRecieved + " " + revIDReceived, logFilePath);
                else
                    Utility.Log("Item id and rev id to process is not received. " + itemIDRecieved + " " + revIDReceived, logFilePath);
            }


            try
            {
                getSourceFileCount(stageDir);
            }
            catch (Exception ex)
            {
                Utility.Log("Could not get total source file count" + Environment.NewLine + ex.ToString(), logFilePath);
            }

            Thread progressThread = null;
            try
            {
                progressThread = new Thread(() => calculateProgress(stageDir));
                progressThread.Start();
            }
            catch (Exception ex)
            {
                Utility.Log("Could not start thread" + Environment.NewLine + ex.ToString(), logFilePath);
            }

            try
            { 
                // Sindhuja | 23/05/2025 | If CreatePDF is True, convert DFT to PDF
                // Sindhuja | 23/05/2025 | If CreateDXF is True, convert PSM to DXF
                if (createPDF == true)
                {
                    // PDF
                    LTC_PROCESS_DRAFT.PROCESS_DRAFT_2(stageDir, outPutFolderPath, logFilePath);
                    
                }
                if(createDXF==true)
                {
                    LTC_PROCESS_PSM.PROCESS_SHEETMETAL_To_DXF(stageDir, outPutFolderPath, logFilePath);
                }
                else
                {
                    Utility.Log("createPDF and createDXF is set to FALSE..", failureFilePath);
                }
            }
            catch (Exception ex)
            {
                Utility.Log("Error encountered while processing drafts. Contact PLM-S", failureFilePath );
                Console.WriteLine("PROCESS_DRAFT_2: " + ex.Message);
                Console.WriteLine("PROCESS_DRAFT_2: " + ex.StackTrace);
            }

            // STP

            if (createSTP == true)
            {
                SOL_PROCESS_PART.PROCESS_PART(stageDir, outPutFolderPath, logFilePath);
                SOL_PROCESS_ASM.PROCESS_ASM(stageDir, outPutFolderPath, logFilePath);
                LTC_PROCESS_PSM.PROCESS_SHEETMETAL_To_STP(stageDir, outPutFolderPath, logFilePath);
                SOL_PROCESS_WELDMENT.PROCESS_WELDMENT(stageDir, outPutFolderPath, logFilePath);
            }         

            try
            {
                getDeliveralesCount(outPutFolderPath);
            }
            catch (Exception ex)
            {
                Utility.Log("Could not get total deliverable count" + Environment.NewLine + ex.ToString(), logFilePath);
            }


            SOL_STP_UPLOAD.Stp_Upload(outPutFolderPath, stageDir, logFilePath);
            // 28 - SEPT Running PDFWaterMark After the Solid Edge Session is KILLED
            SE_PDF_UPLOAD.Pdf_Upload(outPutFolderPath, logFilePath);
            SE_DXF_UPLOAD.Dxf_Upload(outPutFolderPath, logFilePath);

            LTC_SE_SET_PROPERTIES.releaseCustomPropDictionary();
            if (System.IO.Directory.Exists(networkFolderPath) == true && networkFolderPath != null &&
               networkFolderPath.Equals("") == false)
            {
                SOL_Network_Copy.copyFilesOverNetwork(outPutFolderPath, networkFolderPath, "PDF");
            }

            if (System.IO.Directory.Exists(networkFolderPathForDXFCopy) == true && networkFolderPathForDXFCopy != null &&
               networkFolderPathForDXFCopy.Equals("") == false)
            {
                SOL_Network_Copy.copyFilesOverNetwork(outPutFolderPath, networkFolderPathForDXFCopy, "DXF");
            }

            Utility.Log("Utility Completed @ " + System.DateTime.Now.ToString(), logFilePath);

            try
            {
                if (progressThread != null)
                    progressThread.Abort();
            }
            catch (Exception ex)
            {
                Utility.Log("Could not abort thread" + Environment.NewLine + ex.ToString(), logFilePath);
            }

            try
            {
                Utility.Log("Reached calculate progress", logFilePath);
                Utility.Log("Sleeping for 5 seconds", logFilePath);
                System.Threading.Thread.Sleep(5000);

                double fileProcessedProgress = 0, deliverableProgress = 0, totalProgress = 0;
                Utility.Log("totalFilesProcessed " + totalFilesProcessed.ToString(), logFilePath);
                Utility.Log("totalFileCountInSession " + totalFileCountInSession.ToString(), logFilePath);
                if (totalFilesProcessed != 0 & totalFileCountInSession != 0)
                    fileProcessedProgress = (totalFilesProcessed / totalFileCountInSession) * 100;
                Utility.Log("fileProcessedProgress " + fileProcessedProgress.ToString(), logFilePath);

                Utility.Log("totalDeliverablesUploaded " + totalDeliverablesUploaded.ToString(), logFilePath);
                Utility.Log("totalDeliverablesCreated " + totalDeliverablesCreated.ToString(), logFilePath);
                if (totalDeliverablesUploaded != 0 & totalDeliverablesCreated != 0)
                    deliverableProgress = (totalDeliverablesUploaded / totalDeliverablesCreated) * 100;
                Utility.Log("deliverableProgress " + deliverableProgress.ToString(), logFilePath);

                Utility.Log("calculating progress", logFilePath);
                Utility.Log("fileProcessedProgress " + fileProcessedProgress.ToString(), logFilePath);
                Utility.Log("deliverableProgress " + deliverableProgress.ToString(), logFilePath);
                if (fileProcessedProgress != 0 & deliverableProgress != 0)
                    totalProgress = 20 + Math.Round((((fileProcessedProgress * 0.5) / 100) + ((deliverableProgress * 0.25) / 100)) * 100, 0);
                else if (fileProcessedProgress != 0 & deliverableProgress == 0)
                    totalProgress = 20 + Math.Round(fileProcessedProgress, 0);
                Utility.Log("totalProgress " + totalProgress.ToString(), logFilePath);

                string filePath = Path.Combine(stageDir, "Progress.txt");
                File.WriteAllText(filePath, totalProgress.ToString());
                Utility.Log("progress written to file in stage directory", logFilePath);
            }
            catch (Exception ex)
            {
                Utility.Log("Could not calculate final progress" + Environment.NewLine + ex.ToString(), logFilePath);
            }


            if (File.Exists(failureFilePath))
            {
                string[] allLines = File.ReadAllLines(failureFilePath);
                DirectoryInfo dir_info = new DirectoryInfo(stageDir);
                string dir_Name = dir_info.Name;
                List<string> fileContents = new List<string>();
                fileContents.Add("Task ID " + dir_Name + " failed because of the following reasons");
                fileContents.Add(" ");
                foreach (string s in allLines)
                    fileContents.Add(s);
                File.WriteAllLines(failureFilePath, fileContents.ToArray() );
            }
        }

        public static double totalFileCountInSession = 0;
        public static double totalFilesProcessed = 0;
        public static void getSourceFileCount(string StageDir)
        {
            sw.WriteLine("getSourceFileCount reached");
            int DraftFilesCount = Directory.GetFiles(StageDir, "*")
                                         .Select(path => Path.GetFullPath(path))
                                         .Where(x => (x.ToLower().EndsWith(".dft")))
                                         .ToArray().Length;


            int PartFilesCount = Directory.GetFiles(StageDir, "*")
                                         .Select(path => Path.GetFullPath(path))
                                         .Where(x => (x.ToLower().EndsWith(".par")))
                                         .ToArray().Length;


            int AsmFilesCount = Directory.GetFiles(StageDir, "*")
                                        .Select(path => Path.GetFullPath(path))
                                        .Where(x => (x.ToLower().EndsWith(".asm")))
                                        .ToArray().Length;


            int sheetMetalFilesCount = Directory.GetFiles(StageDir, "*")
                                             .Select(path => Path.GetFullPath(path))
                                             .Where(x => (x.ToLower().EndsWith(".psm")))
                                             .ToArray().Length;


            int weldmentFilesLength = Directory.GetFiles(StageDir, "*")
                                             .Select(path => Path.GetFullPath(path))
                                             .Where(x => (x.ToLower().EndsWith(".pwd")))
                                             .ToArray().Length;

            totalFileCountInSession = DraftFilesCount + PartFilesCount + AsmFilesCount + sheetMetalFilesCount + weldmentFilesLength;
            sw.WriteLine("totalFileCountInSession " + totalFileCountInSession.ToString());
        }

        public static double totalDeliverablesCreated = 0;
        public static double totalDeliverablesUploaded = 0;
        public static void getDeliveralesCount(string folderPath)
        {
            sw.WriteLine("getDeliveralesCount reached");
            int STPFilesCount = Directory.GetFiles(folderPath, "*")
                                         .Select(path => Path.GetFileName(path))
                                         .Where(x => (x.ToLower().EndsWith(".stp")))
                                         .ToArray().Length;

            int PDFFilesCount = Directory.GetFiles(folderPath, "*")
                                         .Select(path => Path.GetFileName(path))
                                         .Where(x => (x.ToLower().EndsWith(".pdf")))
                                         .ToArray().Length;

            int DXFFilesCount = Directory.GetFiles(folderPath, "*")
                                         .Select(path => Path.GetFileName(path))
                                         .Where(x => (x.ToLower().EndsWith(".dxf")))
                                         .ToArray().Length;

            totalDeliverablesCreated = STPFilesCount + PDFFilesCount + DXFFilesCount;
            sw.WriteLine("totalDeliverablesCreated " + totalDeliverablesCreated.ToString());
        }

        public static void calculateProgress(string stageDir)
        {
            for (; ; )
            {
                sw.WriteLine("Reached calculate progress");
                sw.WriteLine("Sleeping for 5 seconds");
                System.Threading.Thread.Sleep(5000);

                double fileProcessedProgress = 0, deliverableProgress = 0, totalProgress = 20;
                sw.WriteLine("totalFilesProcessed " + totalFilesProcessed.ToString());
                sw.WriteLine("totalFileCountInSession " + totalFileCountInSession.ToString());
                if (totalFilesProcessed != 0 & totalFileCountInSession != 0)
                    fileProcessedProgress = (totalFilesProcessed / totalFileCountInSession) * 100;
                sw.WriteLine("fileProcessedProgress " + fileProcessedProgress.ToString());

                sw.WriteLine("totalDeliverablesUploaded " + totalDeliverablesUploaded.ToString());
                sw.WriteLine("totalDeliverablesCreated " + totalDeliverablesCreated.ToString());
                if (totalDeliverablesUploaded != 0 & totalDeliverablesCreated != 0)
                    deliverableProgress = (totalDeliverablesUploaded / totalDeliverablesCreated) * 100;
                sw.WriteLine("deliverableProgress " + deliverableProgress.ToString());

                sw.WriteLine("calculating progress");
                sw.WriteLine("fileProcessedProgress " + fileProcessedProgress.ToString());
                sw.WriteLine("deliverableProgress " + deliverableProgress.ToString());
                if (fileProcessedProgress != 0 & deliverableProgress != 0)
                    totalProgress = 20 + Math.Round((((fileProcessedProgress * 0.5) / 100) + ((deliverableProgress * 0.25) / 100)) * 100, 0);
                else if (fileProcessedProgress != 0 & deliverableProgress == 0)
                    totalProgress = 20 + Math.Round(((fileProcessedProgress * 0.5) / 100) * 100, 0);
                sw.WriteLine("totalProgress " + totalProgress.ToString());

                string filePath = Path.Combine(stageDir, "Progress.txt");
                File.WriteAllText(filePath, totalProgress.ToString());
                sw.WriteLine("progress written to file in stage directory");
            }
        }
    }
}
