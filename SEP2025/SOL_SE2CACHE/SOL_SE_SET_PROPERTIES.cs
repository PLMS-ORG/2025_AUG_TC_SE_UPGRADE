using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SolidEdgeFileProperties;
using SolidEdge;
using System.IO;
using System.Runtime.InteropServices;
using LTC_SE2CACHE;

namespace LTC_SE2CACHE
{
    class LTC_SE_SET_PROPERTIES
    {
        static int itemIDIndex = 1;
        static int revIndex = 2;
        static int fileNameIndex = 0;
        static int cadFileDirectoryIndex = 3;
        static int datasetNameIndex = 4;
        private static int RELEASESTATUSINDEX = 5;
        static int ProjectIndex = 6;

        static String ITEM_ID = "ITEM_ID";
        static String REV_ID = "REV_ID";
        static String FILE_NAME = "FILE_NAME";
        static String CADFILE_DIRECTORY = "FILE_DIRECTORY";
        static String DATASET_NAME = "DATASETNAME";      
        static String PROJECT = "PROJECT";

        static String RELEASE_STATUS = "RELEASE_STATUS";
     

        static Dictionary<String, string[]> AssemcustomPropDictionary = new Dictionary<string, string[]>();
        static Dictionary<String, string[]> PartcustomPropDictionary = new Dictionary<string, string[]>();
        static Dictionary<String, string[]> DraftcustomPropDictionary = new Dictionary<string, string[]>();
        static Dictionary<String, string[]> SheetMetalcustomPropDictionary = new Dictionary<string, string[]>();
        static Dictionary<String, string[]> WeldmentcustomPropDictionary = new Dictionary<string, string[]>();

        static Dictionary<String, String> TypeDictionary = new Dictionary<string, string>();

        private static void readCustomPropertyBasedOndsType(string dirToScan, String dsType)
        {
            string[] bomPropertyFile = Directory.GetFiles(dirToScan, "SOL_props_bom_" + dsType + "*")
                                        .Select(path => Path.GetFullPath(path)) // -- GetFullPath - Works
                                        .Where(x => (x.EndsWith(".txt") || x.EndsWith(".TXT")))
                                        .ToArray();
            if (bomPropertyFile == null || bomPropertyFile.Length == 0)
            {
                Console.WriteLine("LTC_SE_SET_PROPERTIES:- NO " + dsType + " BOM PROPERTY FILE IDENTIFIED");
                return;
            }

            /*strVfileName + TILDE + strVItemID + TILDE + strVRevId + TILDE + strVcadFileDirectory + TILDE + strVDatasetName + TILDE +              Release_Status + TILDE + Project_Ids*/

            List<string[]> LineArr = File.ReadLines(bomPropertyFile[0], Encoding.Default)
                          .Select(line => line.Split('~'))
                          .ToList();
            if (LineArr.Count == 0)
            {
                Console.WriteLine("LTC_SE_SET_PROPERTIES:- NO LINES IN BOM PROPERTY FILE IDENTIFIED: " + bomPropertyFile[0]);
                return;

            }
            foreach (string[] lineR in LineArr)
            {
                String fileName = lineR[0];
                if (fileName != null && fileName.Equals("") == false)
                {
                    // change the fileName to lower case...for comparison later.
                    fileName = fileName.ToLower();
                    //Console.WriteLine("fileName: " + fileName);
                    //Console.WriteLine("lineR: " + lineR[0]);
                    //Console.WriteLine("lineR: " + lineR[1]);
                    //Console.WriteLine("lineR: " + lineR[2]);
                    if (dsType.Equals("ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        if (AssemcustomPropDictionary.ContainsKey(fileName) == false)
                        {
                            AssemcustomPropDictionary.Add(fileName, lineR);
                        }
                        if (TypeDictionary.ContainsKey(lineR[1]) == false)
                        {
                            TypeDictionary.Add(lineR[1], "SE Assembly");
                        }
                    }
                    if (dsType.Equals("PART", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        if (PartcustomPropDictionary.ContainsKey(fileName) == false)
                        {
                            PartcustomPropDictionary.Add(fileName, lineR);
                        }
                        if (TypeDictionary.ContainsKey(lineR[1]) == false)
                        {
                            TypeDictionary.Add(lineR[1], "SE Part");
                        }
                    }
                    if (dsType.Equals("DRAFT", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        if (DraftcustomPropDictionary.ContainsKey(fileName) == false)
                        {
                            DraftcustomPropDictionary.Add(fileName, lineR);
                        }
                    }
                    if (dsType.Equals("SHEETMETAL", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        if (SheetMetalcustomPropDictionary.ContainsKey(fileName) == false)
                        {
                            SheetMetalcustomPropDictionary.Add(fileName, lineR);
                        }
                        if (TypeDictionary.ContainsKey(lineR[1]) == false)
                        {
                            TypeDictionary.Add(lineR[1], "SHEETMETAL");
                        }
                    }
                    if (dsType.Equals("WELDMENT", StringComparison.OrdinalIgnoreCase) == true)
                    {
                        if (WeldmentcustomPropDictionary.ContainsKey(fileName) == false)
                        {
                            WeldmentcustomPropDictionary.Add(fileName, lineR);
                        }
                        if (TypeDictionary.ContainsKey(lineR[1]) == false)
                        {
                            TypeDictionary.Add(lineR[1], "WELDMENT");
                        }
                    }
                }

            }


        }
        public static void readCustomProperty(string dirToScan)
        {
            readCustomPropertyBasedOndsType(dirToScan, "ASSEMBLY");
            readCustomPropertyBasedOndsType(dirToScan, "PART");
            readCustomPropertyBasedOndsType(dirToScan, "DRAFT");
            readCustomPropertyBasedOndsType(dirToScan, "SHEETMETAL");
            readCustomPropertyBasedOndsType(dirToScan, "WELDMENT");
        }

        public static String getParent(String itemID)
        {
            if (TypeDictionary != null && TypeDictionary.Count > 0)
            {
                String parent = "";
                TypeDictionary.TryGetValue(itemID, out parent);
                return parent;

            }
            return "";
        }

        public static String getProperty(String dsType, String fileName, String PropertyName)
        {
            string[] values = null;
            fileName = fileName.ToLower();
            if (dsType.Equals("ASSEMBLY", StringComparison.OrdinalIgnoreCase) == true)
            {
                if (AssemcustomPropDictionary != null && AssemcustomPropDictionary.Count > 0)
                {
                    //Console.WriteLine("fileNAME: " + fileName);                      
                    AssemcustomPropDictionary.TryGetValue(fileName, out values);
                    //Console.WriteLine("lineR: " + values[0]);
                    //Console.WriteLine("lineR: " + values[1]);
                    //Console.WriteLine("lineR: " + values[2]);
                }
            }
            if (dsType.Equals("PART", StringComparison.OrdinalIgnoreCase) == true)
            {
                if (PartcustomPropDictionary != null && PartcustomPropDictionary.Count > 0)
                {
                    PartcustomPropDictionary.TryGetValue(fileName, out values);
                }
            }
            if (dsType.Equals("DRAFT", StringComparison.OrdinalIgnoreCase) == true)
            {
                if (DraftcustomPropDictionary != null && DraftcustomPropDictionary.Count > 0)
                {
                    //Console.WriteLine("fileNAME: " + fileName);  
                    DraftcustomPropDictionary.TryGetValue(fileName, out values);
                    //Console.WriteLine("lineR: " + values[0]);
                    //Console.WriteLine("lineR: " + values[1]);
                    //Console.WriteLine("lineR: " + values[2]);
                }
            }
            if (dsType.Equals("SHEETMETAL", StringComparison.OrdinalIgnoreCase) == true)
            {
                if (SheetMetalcustomPropDictionary != null && SheetMetalcustomPropDictionary.Count > 0)
                {
                    SheetMetalcustomPropDictionary.TryGetValue(fileName, out values);
                }
            }
            if (dsType.Equals("WELDMENT", StringComparison.OrdinalIgnoreCase) == true)
            {
                if (WeldmentcustomPropDictionary != null && WeldmentcustomPropDictionary.Count > 0)
                {
                    WeldmentcustomPropDictionary.TryGetValue(fileName, out values);
                }
            }

            if (values == null || values.Length == 0)
            {
                Console.WriteLine("NO VALUE FOUND FOR : " + fileName);
                return "";
            }
            //Console.WriteLine(values[itemIDIndex]);
            //Console.WriteLine(values[revIndex]);
            if (PropertyName.Equals(ITEM_ID, StringComparison.OrdinalIgnoreCase) == true)
            {
                if (values.Length > 0)
                {
                    return values[itemIDIndex];
                }
                else
                {
                    Console.WriteLine("ITEM_ID VALUE NOT FOUND FOR : " + fileName);
                    return "";
                }
            }

            if (PropertyName.Equals(REV_ID, StringComparison.OrdinalIgnoreCase) == true)
            {
                if (values.Length > 0)
                {
                    return values[revIndex];
                }
                else
                {
                    Console.WriteLine("REV_ID VALUE NOT FOUND FOR : " + fileName);
                    return "";
                }
            }

           
            if (PropertyName.Equals(RELEASE_STATUS, StringComparison.OrdinalIgnoreCase) == true)
            {
                if (values.Length > 0)
                {
                    return values[RELEASESTATUSINDEX];
                }
                else
                {
                    Console.WriteLine("RELEASE_STATUS VALUE NOT FOUND FOR : " + fileName);
                    return "";
                }
            }

            if (PropertyName.Equals(PROJECT, StringComparison.OrdinalIgnoreCase) == true)
            {
                if (values.Length > 0)
                {
                    return values[ProjectIndex];
                }
                else
                {
                    Console.WriteLine("PROJECT VALUE NOT FOUND FOR : " + fileName);
                    return "";
                }
            }
           
            //if (PropertyName.Equals(REVISE_REASON, StringComparison.OrdinalIgnoreCase) == true)
            //{
            //    if (values.Length > 0)
            //    {
            //        return values[REVISEREASONINDEX];
            //    }
            //    else
            //    {
            //        Console.WriteLine("REVISE_REASON VALUE NOT FOUND FOR : " + fileName);
            //        return "";
            //    }
            //}

            //if (PropertyName.Equals(DATASETTYPE, StringComparison.OrdinalIgnoreCase) == true)
            //{
            //    if (values.Length > 0)
            //    {
            //        return values[DATASETTYPEINDEX];
            //    }
            //    else
            //    {
            //        Console.WriteLine("DATASETTYPE VALUE NOT FOUND FOR : " + fileName);
            //        return "";
            //    }
            //}
            return "";

        }






        public static void releaseCustomPropDictionary()
        {
            AssemcustomPropDictionary.Clear();
            PartcustomPropDictionary.Clear();
            DraftcustomPropDictionary.Clear();
            SheetMetalcustomPropDictionary.Clear();
            WeldmentcustomPropDictionary.Clear();
        }

    }


}
