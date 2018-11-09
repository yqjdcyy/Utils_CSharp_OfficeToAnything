using Microsoft.Office.Interop.Excel;
using NLog;
using OfficeToAnything.Error;
using System;
using System.IO;
using System.Runtime.InteropServices;

namespace OfficeToAnything.Utils
{
    public class ExcelUtils
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern System.IntPtr GetWindowThreadProcessId(System.IntPtr hwnd, out int ID);

        static Logger logger = LogManager.GetCurrentClassLogger();

        public static String ConvertToPDF(String filePath, String destPath)
        {
            // check
            if (!filePath.IsNormalized() || !File.Exists(filePath))
                throw new Exception("未指定文件");
            if (!filePath.ToLower().EndsWith("xls") && !filePath.ToLower().EndsWith("xlsx"))
                throw new Exception("文件非 Excel 类型");

            // init
            ApplicationClass excel = new ApplicationClass();
            String destFilePath = "";
            try
            {
                // init
                Workbook book = excel.Workbooks.Open(filePath, ReadOnly: true, Password: "");
                destFilePath = Path.Combine(destPath, Guid.NewGuid().ToString() + ".pdf");

                // cehck
                if (0 >= book.Sheets.Count
                    || (0 == book.Sheets.HPageBreaks.Count && 0 == book.Sheets.VPageBreaks.Count))
                    throw new OperationException("文件无内容，请确认");

                // saveas
                book.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, Filename: destFilePath, IgnorePrintAreas: false, Quality: XlFixedFormatQuality.xlQualityStandard, OpenAfterPublish: false);
                book.Close();
                book = null;
            }
            catch(OperationException e)
            {
                throw e;
            }
            catch (Exception e)
            {
                String msg = e.Message;
                if (msg.Contains("密码") || msg.Contains("无法打开文件"))
                {
                    throw new OpenException(msg);
                }
                Console.WriteLine(e.ToString());
            }
            finally
            {
                System.IntPtr t = new IntPtr(excel.Hwnd);
                int procID = 0;
                GetWindowThreadProcessId(t, out procID);
                excel.Quit();
                if (procID != 0)
                {
                    try
                    {
                        var proc = System.Diagnostics.Process.GetProcessById(procID);
                        proc.Kill();
                    }
                    catch (Exception ex)
                    {
                        logger.Error("can`t quit with process kill:{0}", ex.Message);
                    }
                }
            }

            return destFilePath;
        }
    }
}
