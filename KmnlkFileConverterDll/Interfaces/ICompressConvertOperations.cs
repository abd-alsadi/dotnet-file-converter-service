using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KmnlkFileConverterDll.Interfaces
{
   public interface ICompressConvertOperations
    {
        string convertFolderToZipFile(string dataFolderPath, string path);
        string extractFolderFromZipFile(string dataFolderPath, string path);
        string convertFolderToRarFile(string dataFolderPath, string path);
        string extractFolderFromRarFile(string dataFolderPath, string path);
    }
}
