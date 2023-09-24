using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KmnlkFileConverterDll.Interfaces
{
   public interface IExcelConvertOperations
    {
        string convertExcelToText(string dataFolderPath, string path);
        string convertExcelToPdf(string dataFolderPath, string path);
        string convertExcelToHtml(string dataFolderPath, string path);
        string convertExcelToExcel(string dataFolderPath, string path);
        string convertExcelToWord(string dataFolderPath, string path);
        string convertExcelToXml(string dataFolderPath, string path);
        string convertExcelToRtf(string dataFolderPath, string path);
    }
}
