using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DemoAddInTC.services
{
    
        class TCPropertyReader
        {
            private static Dictionary<String, String> list;
            private static String filename;

            

            public TCPropertyReader(String file)
            {
                reload(file);
            }

            public static String get(String field)
            {
                return (list.ContainsKey(field)) ? (list[field]) : (null);
            }

            public void set(String field, Object value)
            {
                if (!list.ContainsKey(field))
                    list.Add(field, value.ToString());
                else
                    list[field] = value.ToString();
            }

            public void Save()
            {
                Save(filename);
            }

            public void Save(String filename1)
            {
                filename = filename1;

                if (!System.IO.File.Exists(filename))
                    System.IO.File.Create(filename);

                System.IO.StreamWriter file = new System.IO.StreamWriter(filename);

                foreach (String prop in list.Keys.ToArray())
                    if (!String.IsNullOrWhiteSpace(list[prop]))
                        file.WriteLine(prop + "=" + list[prop]);

                file.Close();
            }

            public void reload()
            {
                reload(filename);
            }

            public static void reload(String filename1)
            {
                filename = filename1;
                list = new Dictionary<String, String>();

                if (System.IO.File.Exists(filename))
                    loadFromFile(filename);
                else
                    System.IO.File.Create(filename);
            }

            public static void loadFromFile(String file)
            {
                foreach (String line in System.IO.File.ReadAllLines(file))
                {
                    if ((!String.IsNullOrEmpty(line)) &&
                        (!line.StartsWith(";")) &&
                        (!line.StartsWith("#")) &&
                        (!line.StartsWith("'")) &&
                        (line.Contains('=')))
                    {
                        int index = line.IndexOf('=');
                        String key = line.Substring(0, index).Trim();
                        String value = line.Substring(index + 1).Trim();

                        if ((value.StartsWith("\"") && value.EndsWith("\"")) ||
                            (value.StartsWith("'") && value.EndsWith("'")))
                        {
                            value = value.Substring(1, value.Length - 2);
                        }

                        try
                        {
                            //ignore dublicates
                            list.Add(key, value);
                        }
                        catch (Exception ex)
                        {
                            
                        }
                    }
                }
            }
       
    }
}
