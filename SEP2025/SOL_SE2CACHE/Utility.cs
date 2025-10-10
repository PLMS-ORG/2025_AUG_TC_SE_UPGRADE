using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LTC_SE2CACHE
{
    class Utility
    {
        // Sindhuja ||23-06-2025|| Set the ignoreStatus ="YES" || Request from Sanju
        //public static String IgnoreReleaseStatus = "NO";
        //public static String IgnoreReleaseStatus = "YES";

        public static String GetClickOnceLocation()
        {
            //Get the assembly information
            System.Reflection.Assembly assemblyInfo = System.Reflection.Assembly.GetExecutingAssembly();
            //Location is where the assembly is run from 
            string assemblyLocation = assemblyInfo.Location;
            //CodeBase is the location of the ClickOnce deployment files
            Uri uriCodeBase = new Uri(assemblyInfo.CodeBase);
            string ClickOnceLocation = Path.GetDirectoryName(uriCodeBase.LocalPath.ToString());
            return ClickOnceLocation;

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

        public static String getMatchingFile(String fileName, String extension, String dirToSearch)
        {
            FileInfo f = new FileInfo(fileName);

            if (f.Extension.Equals(".stp", StringComparison.OrdinalIgnoreCase) == true)
            {
                DirectoryInfo DirToSearch = new DirectoryInfo(f.DirectoryName);
                FileInfo[] filesInDir = DirToSearch.GetFiles(Path.GetFileNameWithoutExtension(fileName) + ".par");
                Console.WriteLine("Identified Number of PAR: " + filesInDir.Length);
                if (filesInDir.Length == 1)
                {
                    return ".par";                    
                }

                filesInDir = DirToSearch.GetFiles(Path.GetFileNameWithoutExtension(fileName) + ".asm");
                Console.WriteLine("Identified Number of ASM: " + filesInDir.Length);
                if (filesInDir.Length == 1)
                {
                    return ".asm";
                }

                filesInDir = DirToSearch.GetFiles(Path.GetFileNameWithoutExtension(fileName) + ".psm");
                Console.WriteLine("Identified Number of PSM: " + filesInDir.Length);
                if (filesInDir.Length == 1)
                {
                    return ".psm";
                }

                filesInDir = DirToSearch.GetFiles(Path.GetFileNameWithoutExtension(fileName) + ".pwd");
                Console.WriteLine("Identified Number of PWD: " + filesInDir.Length);
                if (filesInDir.Length == 1)
                {
                    return ".pwd";
                }

                return "";

            }
            return "";
        }
       

    }
}
