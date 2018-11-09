using Microsoft.Office.Interop.PowerPoint;
using OfficeToAnything.Error;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;

namespace OfficeToAnything.Utils
{
    public class PPTUtils
    {
        public static String ConvertToPDF(String filePath, String destPath)
        {
            // file exist & check type
            if (!filePath.IsNormalized() || !File.Exists(filePath))
                throw new Exception("未指定文件");
            if (!filePath.ToLower().EndsWith("ppt") && !filePath.ToLower().EndsWith("pptx"))
                throw new Exception("文件非 Powerpoint 类型");

            Application app = new Application();
            String destFilePath = String.Empty;
            try
            {
                Presentation presentation =
                    app.Presentations.Open2007(
                    filePath,
                    Microsoft.Office.Core.MsoTriState.msoCTrue,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoFalse);

                destFilePath = Path.Combine(destPath, Guid.NewGuid().ToString() + ".pdf");
                presentation.SaveAs(destFilePath, PpSaveAsFileType.ppSaveAsPDF);
                presentation.Close();

            }
            catch (Exception e)
            {
                String method = (null != e.TargetSite) ? e.TargetSite.Name : "";
                switch (method)
                {
                    case "Open2007":
                        throw new OpenException(e.Message);
                    case "SaveAs":
                        throw new OperationException(e.Message);
                    default:
                        Console.WriteLine(e.ToString());
                        break;
                }
            }
            finally
            {
                app.Quit();
            }

            return destFilePath;
        }

        public static List<String> ConvertToIMAGE(String filePath, String destPath)
        {
            // init
            Application app = new Application();
            List<String> list = new List<string>();

            try
            {
                // init
                Presentation presentation = app.Presentations.Open2007(
                    filePath,
                    Microsoft.Office.Core.MsoTriState.msoCTrue,
                    Microsoft.Office.Core.MsoTriState.msoFalse, 
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoFalse);
                destPath = Path.Combine(destPath, Guid.NewGuid().ToString());

                // save
                presentation.SaveAs(destPath, PpSaveAsFileType.ppSaveAsJPG);
                presentation.Close();
                app.Quit();

                // rename
                String[] fileNames = Directory.GetFiles(destPath);
                Regex regex = new Regex("\\d+");
                foreach (var fileName in fileNames)
                {
                    String newFileName = Path.Combine(Path.GetDirectoryName(fileName), regex.Match(Path.GetFileNameWithoutExtension(fileName)).Value + Path.GetExtension(fileName).ToLower());
                    File.Move(fileName, newFileName);
                    list.Add(newFileName);
                }
            }
            catch (Exception e)
            {
                String method = (null != e.TargetSite) ? e.TargetSite.Name : "";
                switch (method)
                {
                    case "Open2007":
                        throw new OpenException(e.Message);
                    case "SaveAs":
                        throw new OperationException(e.Message);
                    default:
                        Console.WriteLine(e.ToString());
                        break;
                }
            }

            return list;
        }
    }
}
