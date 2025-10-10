using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LTC_SE2CACHE
{
    class SOL_Network_Copy
    {

        public static void copyFilesOverNetwork(String sourceFolder, String targetNetworkFolder, String type)
        {
            if (type.Equals("STP", StringComparison.OrdinalIgnoreCase) == true)
            {
                CopyFilesOverNetwork(sourceFolder, targetNetworkFolder, ".stp");
            }
            if (type.Equals("PDF", StringComparison.OrdinalIgnoreCase) == true)
            {
                CopyFilesOverNetwork(sourceFolder, targetNetworkFolder, ".pdf");
            }

            if (type.Equals("DXF", StringComparison.OrdinalIgnoreCase) == true)
            {
                CopyFilesOverNetwork(sourceFolder, targetNetworkFolder, ".dxf");
            }

        }
        public static void CopyFilesOverNetwork(String sourceFolder, String targetNetworkFolder,String extn)
        {

            string[] STPFiles = Directory.GetFiles(sourceFolder, "*")
                                         .Select(path => Path.GetFileName(path))
                                         .Where(x => (x.EndsWith(extn) || x.EndsWith(extn.ToUpper())))
                                         .ToArray();
            if (STPFiles == null || STPFiles.Length == 0)
            {
                Console.WriteLine("SOL_Network_Copy:- NO OF " + extn + " FILES IDENTIFIED");
                return;
            }

            Console.WriteLine("SOL_Network_Copy:- NO OF " + extn + " FILES IDENTIFIED: " + STPFiles.Length);

            foreach (String stpFile in STPFiles)
            {
                String targetStpFile = Path.Combine(targetNetworkFolder, stpFile);

                String sourceStpFile = Path.Combine(sourceFolder, stpFile);
                try
                {
                    File.Copy(sourceStpFile, targetStpFile,true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("SOL_Network_Copy:- File.Copy FAILED: " + sourceStpFile + ":" + targetStpFile);
                    continue;
                }
            }


        }
    }
}
