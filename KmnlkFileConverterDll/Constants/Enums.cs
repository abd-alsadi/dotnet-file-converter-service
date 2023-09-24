using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KmnlkFileConverterDll.Constants
{
    public static class Enums
    {

        public enum Enum_Convert_Result
        {
            SUCCESS,
            NOT_SUCCESS,
            ERROR_PARAMETERS
        }

        public enum Enum_Convert_Type
        {
            WORD,
            PDF,
            EXCEL,
            TEXT,
            HTML,
            XML,
            RTF,
            Zip_File,
            RAR_File,
            Zip_Folder,
            Rar_Folder
        }
    }
}
