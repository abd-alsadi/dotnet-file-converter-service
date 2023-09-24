using KmnlkFileConverterDll.Helpers;
using KmnlkFileConverterDll.Management;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KmnlkCommon.Shareds.LoggerManagement;

namespace Mytest
{
    class Program
    {
        static void Main(string[] args)
        {
            ILog logger = new FileLogger("");
            BussinessFileConvertManagement bb = new BussinessFileConvertManagement(logger);
            string dataFolderPath = @"E:\my projects\KmnlkFileConverter\KmnlkFileConverterApi\DataFolder\pdf";

            string a = bb.convertPdfTo(dataFolderPath, Path.Combine(dataFolderPath, "test1.pdf"), 2);
            //string aa = bb.convertExcelTo(dataFolderPath, Path.Combine(dataFolderPath, "test1.pdf"), 1);
            //string aaa = bb.convertExcelTo(dataFolderPath, Path.Combine(dataFolderPath, "test1.pdf"), 2);
            //string aaaa = bb.convertExcelTo(dataFolderPath, Path.Combine(dataFolderPath, "test1.pdf"), 3);
            //string aaaaa = bb.convertExcelTo(dataFolderPath, Path.Combine(dataFolderPath, "test1.pdf"), 4);
            //string aaaaaa = bb.convertExcelTo(dataFolderPath, Path.Combine(dataFolderPath, "test1.pdf"), 5);
            //string aaaaaaa = bb.convertExcelTo(dataFolderPath, Path.Combine(dataFolderPath, "test1.pdf"), 6);




        }
        public static string convertWord(string dataFolderPath, string path)
        {
            try
            {

               Microsoft.Office.Interop.Word.Application wordManager = new Microsoft.Office.Interop.Word.Application();
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".docx";
                Microsoft.Office.Interop.Word.Document wordDocument = wordManager.Documents.Open(path);
                //wordDocument.ExportAsFixedFormat(newPath, WdExportFormat.wdExportFormatPDF);
                wordDocument.SaveAs2(newPath, WdSaveFormat.wdFormatDocument);
                wordDocument.Close(false, false, false);
                wordManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordManager);
                return newPath;
            }
            catch (Exception e)
            {
                return null;
            }
        }


    }
}
