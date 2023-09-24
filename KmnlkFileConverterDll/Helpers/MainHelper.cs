
using KmnlkFileConverterDll.Constants;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KmnlkFileConverterDll.Constants.Enums;

namespace KmnlkFileConverterDll.Helpers
{
   public class MainHelper
    {
     
        public static string getPathWithOutExt(string path)
        {
            int ind = path.IndexOf(".");
            string res = "";
            if (ind>0)
             res = path.Substring(0, ind);
            return res;
        }
        public static string getStringTypeExt(int type)
        {
            switch ((Enum_Convert_Type)type)
            {
                case Enum_Convert_Type.WORD:
                    return "docx";
                case Enum_Convert_Type.EXCEL:
                    return "xlsx";
                case Enum_Convert_Type.TEXT:
                    return "txt";
                case Enum_Convert_Type.HTML:
                    return "html";
                case Enum_Convert_Type.XML:
                    return "xml";
                case Enum_Convert_Type.RTF:
                    return "rtf";
                case Enum_Convert_Type.PDF:
                    return "pdf";

                case Enum_Convert_Type.Zip_File:
                    return "zip";
                case Enum_Convert_Type.RAR_File:
                    return "rar";

                case Enum_Convert_Type.Zip_Folder:
                    return "folder";
                case Enum_Convert_Type.Rar_Folder:
                    return "folder";


                default:
                    return "pdf";
            }
        }
        //public static int getTextWidth(string text)
        //{
        //    text = text.Trim();
        //    int countChar = 0;
        //    int countDegit = 0;
        //    char[] arr = text.ToArray();
        //    foreach(char ch in arr)
        //    {
        //        if (Char.IsDigit(ch))
        //            countDegit++;
        //        else
        //            countChar++;
        //    }

        //    int textLength = 0;
        //    textLength = (int)(((text.Length * 11) + 35) * 1.2);
        //    if (textLength < 120)
        //        textLength = 120;

        //    return textLength;
        //}
    }
}
