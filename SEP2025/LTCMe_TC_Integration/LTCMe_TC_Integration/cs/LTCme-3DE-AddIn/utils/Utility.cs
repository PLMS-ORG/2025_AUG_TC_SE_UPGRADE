using DemoAddInTC.se;
using DemoAddInTC.services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DemoAddInTC.utils
{
    class Utility
    {
        public static String DE = "creo_data_preparation";

        public static bool downloadFolderExists(String sdExpFile, String logFilePath)
        {

            String stageDir = System.IO.Path.GetDirectoryName(sdExpFile);
            String folderName_sdExp = System.IO.Path.GetFileNameWithoutExtension(sdExpFile);
            String creoHome = Path.Combine(stageDir, folderName_sdExp);
            if (Directory.Exists(creoHome) == false)
            {
                return false;
            }
            else
            {
                Utility.Log(sdExpFile + " Exists Already..", logFilePath);
                return true;
            }

        }

        public static bool CreateDownloadFolderMM(String sdExpFile, String logFilePath)
        {
            String stageDir = System.IO.Path.GetDirectoryName(sdExpFile);
            String folderName_sdExp = System.IO.Path.GetFileNameWithoutExtension(sdExpFile);

            {
                String creoHome = Path.Combine(stageDir, folderName_sdExp);
                if (Directory.Exists(creoHome) == false)
                {
                    Directory.CreateDirectory(creoHome);
                }
                else
                {
                    try
                    {
                        // delete
                        Directory.Delete(creoHome);
                        Directory.CreateDirectory(creoHome);
                    }
                    catch (Exception ex)
                    {
                        Log("Exception in Delete: " + ex.Message, logFilePath);
                        return false;
                    }
                }


                CreoUtilitySession.DownloadFolder = creoHome;
                //Log("CreoUtilitySession.DownloadFolder: " + CreoUtilitySession.DownloadFolder, logFilePath);
                return true;
            }
        }

        public static void DeleteDirectory(string target_dir)
        {
            string[] files = Directory.GetFiles(target_dir);
            string[] dirs = Directory.GetDirectories(target_dir);

            foreach (string file in files)
            {
                File.SetAttributes(file, FileAttributes.Normal);
                File.Delete(file);
            }

            foreach (string dir in dirs)
            {
                DeleteDirectory(dir);
            }

            Directory.Delete(target_dir, false);
        }


        public static String CreateDownloadFolder(String stageDir, String logFilePath, String itemID)
        {

            {
                String creoHome = Path.Combine(stageDir, itemID);
                if (Directory.Exists(creoHome) == false)
                {
                    Directory.CreateDirectory(creoHome);
                }
                else
                {
                    try
                    {

                        DeleteDirectory(creoHome);
                        Directory.CreateDirectory(creoHome);
                    }
                    catch (Exception ex)
                    {
                        Log("Exception in Delete: " + ex.Message, logFilePath);
                        return "";
                    }
                }


                CreoUtilitySession.DownloadFolder = creoHome;
                //Log("CreoUtilitySession.DownloadFolder: " + CreoUtilitySession.DownloadFolder, logFilePath);
                return creoHome;
            }
        }

        public static void Log(string logMessage, string logFilePath)
        {
            try
            {
                StreamWriter w = File.AppendText(logFilePath);
                w.WriteLine("{0}", logMessage);
                Console.WriteLine(logMessage);
                w.Close();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Log Writing Exception: " + ex.Message);
            }

        }

    }
}
