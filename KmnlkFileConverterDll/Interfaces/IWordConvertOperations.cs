using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KmnlkFileConverterDll.Interfaces
{
   public interface IWordConvertOperations
    {
        string convertWordToText(string dataFolderPath, string path);
        string convertWordToPdf(string dataFolderPath, string path);
        string convertWordToHtml(string dataFolderPath, string path);
        string convertWordToWord(string dataFolderPath, string path);
        string convertWordToExcel(string dataFolderPath, string path);
        string convertWordToXml(string dataFolderPath, string path);
        string convertWordToRtf(string dataFolderPath, string path);
    }
}
