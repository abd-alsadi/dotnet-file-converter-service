using KmnlkFileConverterDll.Management;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;
using static KmnlkFileConverterDll.Constants.Enums;
using static KmnlkCommon.Shareds.LoggerManagement;

namespace KmnlkFileConverterDll.Management
{
    public class BussinessFileConvertManagement
    {
        private WordConvertManagement WCM;
        private ExcelConvertManagement ECM;
        private PdfConvertManagement PCM;
        private CompressConvertManagement Zp;
        public ILog logger;
        public BussinessFileConvertManagement(ILog logger)
        {
            this.logger = logger;
            WCM = new WordConvertManagement(logger);
            ECM = new ExcelConvertManagement(logger);
            PCM = new PdfConvertManagement(logger);
            Zp = new CompressConvertManagement(logger);
        }

        public string convertWordTo(string dataFolderPath, string path, int type)
        {
                switch (type)
                {
                    case (int)Enum_Convert_Type.TEXT:
                            return WCM.convertWordToText(dataFolderPath,path);
                    case (int)Enum_Convert_Type.PDF:
                            return WCM.convertWordToPdf(dataFolderPath, path);
                    case (int)Enum_Convert_Type.WORD:
                            return WCM.convertWordToWord(dataFolderPath, path);
                    case (int)Enum_Convert_Type.EXCEL:
                            return WCM.convertWordToExcel(dataFolderPath, path);
                    case (int)Enum_Convert_Type.HTML:
                            return WCM.convertWordToHtml(dataFolderPath, path);
                    case (int)Enum_Convert_Type.XML:
                            return WCM.convertWordToXml(dataFolderPath, path);
                    case (int)Enum_Convert_Type.RTF:
                            return WCM.convertWordToRtf(dataFolderPath, path);
                    default:
                            return WCM.convertWordToText(dataFolderPath, path);
            }
        }

        public string convertExcelTo(string dataFolderPath, string path, int type)
        {
            switch (type)
            {
                case (int)Enum_Convert_Type.TEXT:
                    return ECM.convertExcelToText(dataFolderPath, path);
                case (int)Enum_Convert_Type.PDF:
                    return ECM.convertExcelToPdf(dataFolderPath, path);
                case (int)Enum_Convert_Type.WORD:
                    return ECM.convertExcelToWord(dataFolderPath, path);
                case (int)Enum_Convert_Type.EXCEL:
                    return ECM.convertExcelToExcel(dataFolderPath, path);
                case (int)Enum_Convert_Type.HTML:
                    return ECM.convertExcelToHtml(dataFolderPath, path);
                case (int)Enum_Convert_Type.XML:
                    return ECM.convertExcelToXml(dataFolderPath, path);
                case (int)Enum_Convert_Type.RTF:
                    return ECM.convertExcelToRtf(dataFolderPath, path);
                default:
                    return ECM.convertExcelToText(dataFolderPath, path);
            }
        }

        public string convertPdfTo(string dataFolderPath, string path, int type)
        {
            switch (type)
            {
                case (int)Enum_Convert_Type.TEXT:
                    return PCM.convertPdfToText(dataFolderPath, path);
                case (int)Enum_Convert_Type.PDF:
                    return PCM.convertPdfToPdf(dataFolderPath, path);
                case (int)Enum_Convert_Type.WORD:
                    return PCM.convertPdfToWord(dataFolderPath, path);
                case (int)Enum_Convert_Type.EXCEL:
                    return PCM.convertPdfToExcel(dataFolderPath, path);
                case (int)Enum_Convert_Type.HTML:
                    return PCM.convertPdfToHtml(dataFolderPath, path);
                case (int)Enum_Convert_Type.XML:
                    return PCM.convertPdfToXml(dataFolderPath, path);
                case (int)Enum_Convert_Type.RTF:
                    return PCM.convertPdfToRtf(dataFolderPath, path);
                default:
                    return PCM.convertPdfToText(dataFolderPath, path);
            }
        }

        public string convertCompressOrFolderTo(string dataFolderPath, string pathSource, int type)
        {
            switch (type)
            {
                case (int)Enum_Convert_Type.Zip_File:
                    return Zp.convertFolderToZipFile(dataFolderPath, pathSource);
                case (int)Enum_Convert_Type.RAR_File:
                    return Zp.convertFolderToRarFile(dataFolderPath, pathSource);
                case (int)Enum_Convert_Type.Zip_Folder:
                    return Zp.extractFolderFromZipFile(dataFolderPath, pathSource);
                case (int)Enum_Convert_Type.Rar_Folder:
                    return Zp.extractFolderFromRarFile(dataFolderPath, pathSource);
            }
            return null;
        }
    }
}