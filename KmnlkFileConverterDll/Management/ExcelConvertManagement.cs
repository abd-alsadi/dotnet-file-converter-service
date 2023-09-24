
using KmnlkFileConverterDll.Constants;
using KmnlkFileConverterDll.Exceptions;
using KmnlkFileConverterDll.Helpers;
using KmnlkFileConverterDll.Interfaces;
using KmnlkCommon.Shareds;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KmnlkCommon.Shareds.LoggerManagement;
using Microsoft.Office.Interop.Word;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace KmnlkFileConverterDll.Management
{
    public class ExcelConvertManagement : IExcelConvertOperations, IValidationOperations
    {
        private ILog logger;
        private Microsoft.Office.Interop.Excel.Application excelManager;
        private Microsoft.Office.Interop.Excel.Workbook excelBook;
        public ExcelConvertManagement(ILog logger)
        {
            this.logger = logger;
            excelManager = new Microsoft.Office.Interop.Excel.Application();
        }

        public string convertExcelToExcel(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                excelManager = new Microsoft.Office.Interop.Excel.Application()
                {
                    ScreenUpdating = false,
                    DisplayAlerts = false
                };
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".csv";
                excelBook = excelManager.Workbooks.Open(path, 0, true, 5, "", "", true);
                excelBook.SaveAs(newPath, XlFileFormat.xlCSV);
                //Microsoft.Office.Interop.Excel.Sheets sheets = excelBook.Worksheets;
                //foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
                //{
                //    sheet.ExportAsFixedFormat(
                //           Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                //           newPath,
                //           Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard,
                //           true,
                //           true,
                //           1,
                //           10,
                //           false);
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                //}
                excelBook.Close(false, false, false);
                excelManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (excelBook != null) { excelBook.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook); }
                if (excelManager != null) { excelManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string convertExcelToPdf(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                excelManager = new Microsoft.Office.Interop.Excel.Application() {
                    ScreenUpdating = false,
                    DisplayAlerts = false
                };
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".pdf";
                excelBook = excelManager.Workbooks.Open(path, 0, true, 5, "", "", true);
                excelBook.ExportAsFixedFormat(
                       Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                       newPath,
                       Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard,
                       true,
                       true,
                       1,
                       10,
                       false);
                //Microsoft.Office.Interop.Excel.Sheets sheets = excelBook.Worksheets;
                //foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
                //{
                //    sheet.ExportAsFixedFormat(
                //           Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                //           newPath,
                //           Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard,
                //           true,
                //           true,
                //           1,
                //           10,
                //           false);
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                //}
            excelBook.Close(false, false, false);
                excelManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (excelBook != null) { excelBook.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook); }
                if (excelManager != null) { excelManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string convertExcelToRtf(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                excelManager = new Microsoft.Office.Interop.Excel.Application()
                {
                    ScreenUpdating = false,
                    DisplayAlerts = false
                };
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".csv";
                excelBook = excelManager.Workbooks.Open(path, 0, true, 5, "", "", true);
                excelBook.SaveAs(newPath, XlFileFormat.xlCSV);
                //Microsoft.Office.Interop.Excel.Sheets sheets = excelBook.Worksheets;
                //foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
                //{
                //    sheet.ExportAsFixedFormat(
                //           Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                //           newPath,
                //           Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard,
                //           true,
                //           true,
                //           1,
                //           10,
                //           false);
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                //}
                excelBook.Close(false, false, false);
                excelManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (excelBook != null) { excelBook.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook); }
                if (excelManager != null) { excelManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string convertExcelToText(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                excelManager = new Microsoft.Office.Interop.Excel.Application()
                {
                    ScreenUpdating = false,
                    DisplayAlerts = false
                };
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".txt";
                excelBook = excelManager.Workbooks.Open(path, 0, true, 5, "", "", true);
                excelBook.SaveAs(newPath, XlFileFormat.xlTextWindows);
                //Microsoft.Office.Interop.Excel.Sheets sheets = excelBook.Worksheets;
                //foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
                //{
                //    sheet.ExportAsFixedFormat(
                //           Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                //           newPath,
                //           Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard,
                //           true,
                //           true,
                //           1,
                //           10,
                //           false);
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                //}
                excelBook.Close(false, false, false);
                excelManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (excelBook != null) { excelBook.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook); }
                if (excelManager != null) { excelManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string convertExcelToWord(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                excelManager = new Microsoft.Office.Interop.Excel.Application()
                {
                    ScreenUpdating = false,
                    DisplayAlerts = false
                };
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".csv";
                excelBook = excelManager.Workbooks.Open(path, 0, true, 5, "", "", true);
                excelBook.SaveAs(newPath, XlFileFormat.xlCSV);
                //Microsoft.Office.Interop.Excel.Sheets sheets = excelBook.Worksheets;
                //foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
                //{
                //    sheet.ExportAsFixedFormat(
                //           Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                //           newPath,
                //           Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard,
                //           true,
                //           true,
                //           1,
                //           10,
                //           false);
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                //}
                excelBook.Close(false, false, false);
                excelManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (excelBook != null) { excelBook.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook); }
                if (excelManager != null) { excelManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string convertExcelToXml(string dataFolderPath, string path)
        {

            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                excelManager = new Microsoft.Office.Interop.Excel.Application()
                {
                    ScreenUpdating = false,
                    DisplayAlerts = false
                };
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".xml";
                excelBook = excelManager.Workbooks.Open(path, 0, true, 5, "", "", true);
                excelBook.SaveAs(newPath, XlFileFormat.xlOpenXMLWorkbook);
                //Microsoft.Office.Interop.Excel.Sheets sheets = excelBook.Worksheets;
                //foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
                //{
                //    sheet.ExportAsFixedFormat(
                //           Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                //           newPath,
                //           Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard,
                //           true,
                //           true,
                //           1,
                //           10,
                //           false);
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                //}
                excelBook.Close(false, false, false);
                excelManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (excelBook != null) { excelBook.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook); }
                if (excelManager != null) { excelManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string convertExcelToHtml(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                excelManager = new Microsoft.Office.Interop.Excel.Application()
                {
                    ScreenUpdating = false,
                    DisplayAlerts = false
                };
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".html";
                excelBook = excelManager.Workbooks.Open(path, 0, true, 5, "", "", true);
                excelBook.SaveAs(newPath, XlFileFormat.xlHtml);
                //Microsoft.Office.Interop.Excel.Sheets sheets = excelBook.Worksheets;
                //foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in sheets)
                //{
                //    sheet.ExportAsFixedFormat(
                //           Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                //           newPath,
                //           Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard,
                //           true,
                //           true,
                //           1,
                //           10,
                //           false);
                //    System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                //}
                excelBook.Close(false, false, false);
                excelManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (excelBook != null) { excelBook.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook); }
                if (excelManager != null) { excelManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(excelManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }


        public bool isValid(object model)
        {
            bool result = true;
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
            return result;
        }
    }
}
