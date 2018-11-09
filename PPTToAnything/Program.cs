using OfficeToAnything.Error;
using OfficeToAnything.Utils;
using System;
using System.Collections.Generic;

namespace PPTToAnything
{
    class Program
    {
        static void Main(string[] args)
        {
            //PPT2IMAGE();
            //EXCEL2IMAGE();
            WORD2IMAGE();

            Console.ReadLine();
        }

        private static void WORD2IMAGE()
        {
            handle("empty.doc");
            handle("error-jpg.doc");
            handle("error-ppt.doc");
            handle("normal.doc");
            handle("password-read.doc");
            handle("password-write.doc");
            handle("readonly.doc");

            handle("empty.docx");
            handle("error-jpg.docx");
            handle("error-ppt.docx");
            handle("normal.docx");
            handle("password-read.docx");
            handle("password-writel.docx");
            handle("readonly.docx");

        }
        private static void EXCEL2IMAGE()
        {
            handle("empty.xls");
            handle("error-jpg.xls");
            handle("normal.xls");
            handle("password-read.xls");
            handle("password-write.xls");

            handle("empty.xlsx");
            handle("error-jpg.xlsx");
            handle("error-no-data.xlsx");
            handle("normal.xlsx");
            handle("password-read.xlsx");
            handle("password-write.xlsx");

        }
        private static void PPT2IMAGE()
        {
            handle("empty.ppt");
            handle("error-jpg.ppt");
            handle("error-keynote.ppt");
            handle("error-xls.ppt");
            handle("fix.ppt");
            handle("normal.ppt");
            handle("password-read.ppt");
            handle("password-write.ppt");
            handle("readonly.ppt");

            handle("empty.pptx");
            handle("error-jpg.pptx");
            handle("error-keynote.pptx");
            handle("error-xls.pptx");
            handle("fix.pptx");
            handle("normal.pptx");
            handle("password-read.pptx");
            handle("password-write.pptx");
            handle("readonly.pptx");

        }

        private static String FMT_SUCCESS = "{0} Y\n\t{1}";
        private static String FMT_FAIL= "{0} X\n\t{1}";

        private static void handle(String from)
        {
            // init
            String ext = from.Substring(from.LastIndexOf(".")+ 1).ToLower();
            String to = @"D:\data\export";
            List<String> list = null;
            String path = "";

            // convert
            try
            {
                switch (ext)
                {
                    case "ppt":
                    case "pptx":
                        list = PPTUtils.ConvertToIMAGE(convert(from), to);
                        break;
                    case "xls":
                    case "xlsx":
                        path = ExcelUtils.ConvertToPDF(convert(from), to);
                        break;
                    case "doc":
                    case "docx":
                        path = WordUtils.ConvertToPDF(convert(from), to);
                        break;
                    default:
                        break;
                }
            }
            //catch(OperationException e)
            //{
            //    Console.WriteLine(FMT_FAIL, from, e.Message);
            //    return;
            //}
            //catch (OpenException e)
            //{
            //    Console.WriteLine(FMT_FAIL,from, e.Message);
            //    return;
            //}
            catch (Exception e)
            {
                Console.WriteLine(FMT_FAIL, from, e.ToString());
                return;
            }

            // handle

            switch (ext)
            {
                case "ppt":
                case "pptx":

                    if (null != list && 0 < list.Count)
                    {
                        foreach (String p in list)
                        {
                            Console.WriteLine(p);
                        }
                    }
                    break;
                case "xls":
                case "xlsx":
                case "doc":
                case "docx":
                    Console.WriteLine(FMT_SUCCESS, from, path);
                    break;
                default:
                    break;
            }
            Console.WriteLine();
        }

        private static String convert(String path)
        {
            return System.IO.Path.Combine(
                System.IO.Path.GetDirectoryName(System.Diagnostics.Process.GetCurrentProcess().MainModule.FileName),
                String.Format(
                    "..\\..\\Template\\{0}", 
                    path));
        }
    }
}
