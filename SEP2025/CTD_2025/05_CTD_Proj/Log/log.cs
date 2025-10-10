using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;




namespace Log
{
    class log
    {
        public static string logFile = DateTime.Now.ToString("yyyy_MM_dd_h_mm_ss")+".log";
        public static bool printLogInConsole = false;
        public static void write(logType log_Type,String log_Message)
        {
            string logLine = DateTime.Now.ToString("yyy-MM-dd-h-mm-ss") + " :: " + log_Type.ToString() + " : " + log_Message + "\n";
            System.IO.File.AppendAllText(logFile,logLine );
            Console.WriteLine(logLine);
        }

        public static void writeException( Exception ex,string functionName)
        {
            System.IO.File.AppendAllText(logFile, DateTime.Now.ToString("h-mm-ss") + " :: Exception : "  + ex.Message+".Function name "+functionName + "\n");
            System.IO.File.AppendAllText(logFile, DateTime.Now.ToString("h-mm-ss") + " :: Exception : " + ex.StackTrace + ".Function name " + functionName + "\n");
            System.IO.File.AppendAllText(logFile, DateTime.Now.ToString("h-mm-ss") + " :: Exception : " + ex.InnerException + ".Function name " + functionName + "\n");

            if (printLogInConsole==true)
            {
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.StackTrace);
                Console.WriteLine(ex.InnerException);
            }
        }

    }
}
