using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using iTextSharp.text.pdf;
using iTextSharp.text;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Drawing.Imaging;

namespace LTC_SE2CACHE
{
    class PDFWatermark
    {
        public static bool IsDebug;

        public static void PDFStamp(bool createPDF, String logFilePath,
            String stageDir, String pdfFileNewPath, String projectID)
        {
            if (createPDF == true)
            {
               

                //if (args.Length > 8)
                {
                    bool doStamping = true;
                    string stampingText = null;
                    stampingText = projectID;
                    //for (int i = 8; i < args.Length; i++)
                    //{
                    //    if (args[i].Contains("-stamp="))
                    //    {
                    //        doStamping = true;
                    //        stampingText = args[i];
                    //    }
                    //}


                    if (doStamping == true)
                    {
                        Utility.Log("Stamping required", logFilePath);
                        //string[] stampText = stampingText.Split('=');
                        //if (stampText.Length == 3)
                        {
                            Utility.Log("Stamping will be done as " + stampingText, logFilePath);
                            Utility.Log("Stamping will be done in " + "GREEN" + " color", logFilePath);
                            String UTILUPLOADPATH = "";
                            UTILUPLOADPATH = Utility.GetClickOnceLocation();
                            Console.WriteLine("UTILUPLOADPATH " + UTILUPLOADPATH);
                            Console.WriteLine("Stagig dir " + stageDir);
                            Console.WriteLine("Starting to stamp pdf");
                            if (UTILUPLOADPATH.Equals("") == false)
                            {
                                string[] waterMarkArgs = new string[] { "Text", pdfFileNewPath, Path.Combine(stageDir, "Released.png"), "Times New Roman", "12", "400", stampingText, "White", "GREEN" };
                                PDFWatermark.pdfWaterMarkMain(waterMarkArgs);
                            }
                            Console.WriteLine("Exited pdf stamping");
                        }
                        //else
                          //  Utility.Log("Stamping text could not be obtained from " + stampingText, logFilePath);
                    }
                    else
                        Utility.Log("Stamping not required", logFilePath);
                }
            }


        }
        public static void pdfWaterMarkMain(string[] args)
        {
            try
            {
                foreach (string s in args)
                    Console.WriteLine(s);

                if ((Environment.GetEnvironmentVariable("PDFSTAMPINGDEBUG") == "1" ? true : Environment.GetEnvironmentVariable("PDFSTAMPINGDEBUG") == "ON"))
                {
                    IsDebug = true;
                }
                IsDebug = true; //to remove
                if (!(args[0] == "image"))
                {
                    if (PDFWatermark.IsDebug)
                    {
                        Console.WriteLine("Text Stamping");
                    }
                    Color color = Color.FromName(args[8]);
                    Console.WriteLine("Acquired color");
                    string str = args[2];
                    string str1 = "Times New Roman";
                    Console.WriteLine("Acquired font");
                    if (args[3].Length > 0)
                    {
                        str1 = args[3];
                    }
                    float single = Convert.ToSingle(args[4]);
                    System.Drawing.Font font = new System.Drawing.Font(str1, single, FontStyle.Bold);
                    int num = Convert.ToInt32(args[4]);
                    int num1 = Convert.ToInt32(args[5]);
                    if (IsDebug)
                    {
                        Console.WriteLine("Draw Text");
                    }
                    DrawText(args[6], font, color, num1, str, args[7]);
                    if (PDFWatermark.IsDebug)
                    {
                        Console.WriteLine("Stamping Pdf");
                    }
                    CreateWaterMarkInPDF(args[1], str, num, num1);
                }
                else
                {
                    if (PDFWatermark.IsDebug)
                    {
                        Console.WriteLine("Imange Stamping");
                    }
                    CreateWaterMarkInPDF(args[1], args[2], 300, 300);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            
        }

        public static void DrawText(string text, System.Drawing.Font font, Color textColor, int maxWidth, string path, string background)
        {
            try
            {
                string str = DateTime.Now.ToString("dd/MM/yy");
                System.Drawing.Image bitmap = new Bitmap(1, 1);
                Graphics graphic = Graphics.FromImage(bitmap);
                SizeF sizeF = graphic.MeasureString(string.Concat(text, " ", str), font, maxWidth);
                StringFormat stringFormat = new StringFormat()
                {
                    Trimming = StringTrimming.Word
                };
                bitmap.Dispose();
                graphic.Dispose();
                bitmap = new Bitmap((int)sizeF.Width, (int)sizeF.Height);
                graphic = Graphics.FromImage(bitmap);
                graphic.CompositingQuality = CompositingQuality.HighQuality;
                graphic.InterpolationMode = InterpolationMode.HighQualityBilinear;
                graphic.PixelOffsetMode = PixelOffsetMode.HighQuality;
                graphic.SmoothingMode = SmoothingMode.HighQuality;
                graphic.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
                if (!(background == "transparent"))
                {
                    graphic.Clear(Color.White);
                }
                else
                {
                    graphic.Clear(Color.Transparent);
                }
                Brush solidBrush = new SolidBrush(textColor);
                Console.WriteLine(string.Concat("date to stamp", str));
                graphic.DrawString(string.Concat(text, " ", str), font, solidBrush, new RectangleF(0f, 0f, sizeF.Width, sizeF.Height), stringFormat);
                graphic.Save();
                solidBrush.Dispose();
                graphic.Dispose();
                if (File.Exists(path) == true)
                    File.Delete(path);
                Console.WriteLine("path " + path);
                bitmap.Save(path, ImageFormat.Png);
                bitmap.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

        }

        public static string[] GetAllPdfsToStamp(string path)
        {
            try
            {
                string[] files;
                if (!path.Contains(".pdf"))
                {
                    if (PDFWatermark.IsDebug)
                    {
                        Console.WriteLine("Found Pdf Dir");
                    }
                    files = Directory.GetFiles(path, "*.pdf");
                }
                else
                {
                    if (PDFWatermark.IsDebug)
                    {
                        Console.WriteLine("Found Pdf file");
                    }
                    files = new string[] { path };
                }
                return files;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return null;
            }
            
        }

        public static void CreateWaterMarkInPDF(string PdfFilePath, string stampPath, int height, int width)
        {
            try
            {
                if (PDFWatermark.IsDebug)
                {
                    Console.WriteLine("In CreateWaterMarkInPDF");
                }
                string[] allPdfsToStamp = { PdfFilePath };
                 //PDFWatermark.GetAllPdfsToStamp(PdfFilePath);
                Console.WriteLine("Following pdf's were found");
                foreach (string s in allPdfsToStamp)
                    Console.WriteLine(s);
                int num = 0;
                while (num < (int)allPdfsToStamp.Length)
                {
                    string str = allPdfsToStamp[num];
                    Console.WriteLine("Processing pdf " + str);
                    if ((!File.Exists(stampPath) ? false : File.Exists(str)))
                    {
                        if (PDFWatermark.IsDebug)
                        {
                            Console.WriteLine(string.Concat("Path: ", str));
                        }
                        PdfReader pdfReader = new PdfReader(File.ReadAllBytes(str));
                        //string str1 = str.ToLower().Replace(".pdf", "[temp][file].pdf");
                        PdfStamper pdfStamper = new PdfStamper(pdfReader, new FileStream(str, FileMode.Create));
                        Console.WriteLine("Created temp pdf at " + pdfStamper);
                        if (PDFWatermark.IsDebug)
                        {
                            Console.WriteLine("Created Stamp path");
                        }
                        for (int i = 1; i <= pdfReader.NumberOfPages; i++)
                        {
                            iTextSharp.text.Image instance = iTextSharp.text.Image.GetInstance(stampPath);
                            instance.Alignment = 0;
                            instance.Alignment = 4;
                            instance.ScaleToFit((float)width, (float)height);
                            instance.Alignment = 8;
                            iTextSharp.text.Rectangle pageSize = pdfReader.GetPageSize(i);
                            if (PDFWatermark.IsDebug)
                            {
                                Console.WriteLine(pageSize.Width.ToString());
                            }
                            if (PDFWatermark.IsDebug)
                            {
                                Console.WriteLine(pageSize.Height.ToString());
                            }
                            instance.SetAbsolutePosition(1f, 1f);
                            pdfStamper.GetOverContent(i).AddImage(instance);
                        }
                        if (PDFWatermark.IsDebug)
                        {
                            Console.WriteLine("Created Stamp path");
                        }
                        pdfStamper.FormFlattening = true;
                        pdfStamper.Close();
                        pdfReader.Close();
                        //Console.WriteLine("str1 " + str1);
                        Console.WriteLine("str " + str);
                        //if (File.Exists(str1))
                        //{
                        //    File.Delete(str1);
                        //    File.Move(str1.ToLower().Replace(".pdf", "[temp][file].pdf"), str);
                        //}
                        num++;
                    }
                    else
                    {
                        Console.WriteLine("File " + str + " does not exist ");
                        continue;
                    }
                }
            }
            catch (Exception exception1)
            {
                Exception exception = exception1;
                throw new Exception(string.Concat("Exception: ", exception.Message.ToString()));
            }
            if (PDFWatermark.IsDebug)
            {
                Console.WriteLine("Out CreateWaterMarkInPDF");
            }
        }
    }
}
