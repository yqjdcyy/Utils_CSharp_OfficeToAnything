using Microsoft.Office.Interop.Word;
using OfficeToAnything.Error;
using System;
using System.IO;

namespace OfficeToAnything.Utils
{
    public class WordUtils
    {
        public static String ConvertToPDF(String filePath, String destPath)
        {
            // check
            if (!filePath.IsNormalized() || !File.Exists(filePath))
                throw new Exception("未指定文件");
            if (!filePath.ToLower().EndsWith("doc") && !filePath.ToLower().EndsWith("docx"))
                throw new Exception("文件非 word 类型");

            // init
            ApplicationClass word = new ApplicationClass();
            String destFilePath = "";
            try
            {
                // init
                Document doc = word.Documents.Open(filePath, ReadOnly: true, PasswordDocument: "yunkai", NoEncodingDialog: true, Visible: false, OpenAndRepair: true);
                destFilePath = Path.Combine(destPath, Guid.NewGuid().ToString() + ".pdf");

                // save
                doc.ExportAsFixedFormat(destFilePath, WdExportFormat.wdExportFormatPDF);
                doc.Close(SaveChanges: false);
                doc = null;
            }
            catch (Exception e)
            {
                String method = e.TargetSite.Name;
                switch (method)
                {
                    case "Open":
                    case "OpenNoRepairDialog":
                        throw new OpenException(e.Message);
                }

                Console.WriteLine(e.Message);
            }
            finally
            {
                word.Quit();
                word = null;
            }

            return destFilePath;
        }
    }
}
