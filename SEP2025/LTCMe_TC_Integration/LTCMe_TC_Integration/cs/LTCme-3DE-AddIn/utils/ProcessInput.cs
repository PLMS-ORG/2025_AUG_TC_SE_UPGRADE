using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace DemoAddInTC.utils
{
    class ProcessInput
    {
        public static List<String> readCsv(String inputCsvPath, String logFilePath)
        {
            List<String> itemsList = new List<string>();

            var lines = File.ReadAllLines(inputCsvPath);

            for (var i = 0; i < lines.Length; i += 1)
            {
                Console.WriteLine(lines[i]);
                var line = lines[i];
                
                itemsList.Add(line);

            }

            return itemsList;
        }

    }
}
