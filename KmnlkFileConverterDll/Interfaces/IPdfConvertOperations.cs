using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KmnlkFileConverterDll.Interfaces
{
   public interface IPdfConvertOperations
    {
        string convertPdfToText(string dataFolderPath, string path);
        string convertPdfToPdf(string dataFolderPath, string path);
        string convertPdfToHtml(string dataFolderPath, string path);
        string convertPdfToWord(string dataFolderPath, string path);
        string convertPdfToExcel(string dataFolderPath, string path);
        string convertPdfToXml(string dataFolderPath, string path);
        string convertPdfToRtf(string dataFolderPath, string path);
    }
}
