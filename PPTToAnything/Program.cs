using PPT2HTML5.Expand.Service.CustomException;
using PPT2HTML5.Expand.Service.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PPTToAnything
{
    class Program
    {
        static void Main(string[] args)
        {
            PPT2IMAGE();
            EXCEL2IMAGE();
            WORD2IMAGE();

            Console.ReadLine();
        }

        private static void WORD2IMAGE()
        {
            // empty
            handle("empty.docx");
            // error
            handle("error.docx");
            // password
            handle("password.docx");
            // normal
            handle("one.docx");
        }
        private static void EXCEL2IMAGE()
        {

            // empty
            // empty-not-open   7K
            handle("empty.xlsx");
            // emtpy-just-save  8K
            handle("empty-2.xlsx");
            // error
            handle("error.xlsx");
            // password
            handle("password.xlsx");
            // normal
            handle("one.xlsx");
        }
        private static void PPT2IMAGE()
        {
            // empty
            handle("empty.pptx");
            // readonly
            handle("readonly.pptx");
            // error
            handle("error.pptx");
            // password
            handle("password.pptx");
            // fix
            handle("fix.pptx");
            // normal
            handle("multiple.pptx");
        }

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
                        path= WordUtils.ConvertToPDF(convert(from), to);
                        break;
                    default:
                        break;
                }
            }catch(OperationException e)
            {
                Console.WriteLine("fail to operate file: {0}", e.Message);
                return;
            }
            catch (OpenException e)
            {
                Console.WriteLine("fail to open file: {0}", e.Message);
                return;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
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
                    Console.WriteLine(path);
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
