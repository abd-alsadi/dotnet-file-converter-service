
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
using System.IO;
using Microsoft.Office.Interop.Word;

namespace KmnlkFileConverterDll.Management
{
    public class PdfConvertManagement : IPdfConvertOperations, IValidationOperations
    {
        private ILog logger;
        private Microsoft.Office.Interop.Word.Application PdfManager;
        private Microsoft.Office.Interop.Word.Document PdfDocument;
        public PdfConvertManagement(ILog logger)
        {
            this.logger = logger;
            PdfManager = new Microsoft.Office.Interop.Word.Application();
        }

        public string convertPdfToExcel(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                PdfManager = new Microsoft.Office.Interop.Word.Application();
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".docx";
                PdfDocument = PdfManager.Documents.Open(path);
                PdfDocument.SaveAs2(newPath, WdSaveFormat.wdFormatDocument);
                PdfDocument.Close(false, false, false);
                PdfManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return "";
            }
            catch (Exception e)
            {
            //    if (PdfDocument != null) PdfDocument.Close();
            //    if (PdfManager != null) PdfManager.Quit();
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string convertPdfToHtml(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                PdfManager = new Microsoft.Office.Interop.Word.Application();
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".html";
                PdfDocument = PdfManager.Documents.Open(path);
                PdfDocument.SaveAs2(newPath, WdSaveFormat.wdFormatFilteredHTML);
                PdfDocument.Close(false, false, false);
                PdfManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (PdfDocument != null) { PdfDocument.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfDocument); }
                if (PdfManager != null) { PdfManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string convertPdfToXml(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                PdfManager = new Microsoft.Office.Interop.Word.Application();
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".xml";
                PdfDocument = PdfManager.Documents.Open(path);
                PdfDocument.SaveAs2(newPath, WdSaveFormat.wdFormatXML);
                PdfDocument.Close(false, false, false);
                PdfManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (PdfDocument != null) { PdfDocument.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfDocument); }
                if (PdfManager != null) { PdfManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfManager); } 
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }
        public string convertPdfToRtf(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                PdfManager = new Microsoft.Office.Interop.Word.Application();
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".rtf";
                PdfDocument = PdfManager.Documents.Open(path);
                PdfDocument.SaveAs2(newPath, WdSaveFormat.wdFormatRTF);
                PdfDocument.Close(false, false, false);
                PdfManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (PdfDocument != null) { PdfDocument.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfDocument); }
                if (PdfManager != null) { PdfManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }
        public string convertPdfToPdf(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                PdfManager = new Microsoft.Office.Interop.Word.Application();
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath+ ".pdf";
                 PdfDocument = PdfManager.Documents.Open(path);
                //PdfDocument.ExportAsFixedFormat(newPath, WdExportFormat.wdExportFormatPDF);
                PdfDocument.SaveAs2(newPath, WdSaveFormat.wdFormatPDF);
                PdfDocument.Close(false,false,false);
                PdfManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (PdfDocument != null) { PdfDocument.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfDocument); }
                if (PdfManager != null) { PdfManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string convertPdfToText(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                PdfManager = new Microsoft.Office.Interop.Word.Application();
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".txt";
                PdfDocument = PdfManager.Documents.Open(path);
                PdfDocument.SaveAs2(newPath, WdSaveFormat.wdFormatText);
                PdfDocument.Close(false, false, false);
                PdfManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (PdfDocument != null) { PdfDocument.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfDocument); }
                if (PdfManager != null) { PdfManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string convertPdfToWord(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                PdfManager = new Microsoft.Office.Interop.Word.Application();
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".docx";
                PdfDocument = PdfManager.Documents.Open(path);
                PdfDocument.SaveAs2(newPath, WdSaveFormat.wdFormatDocumentDefault);
                PdfDocument.Close(false, false, false);
                PdfManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (PdfDocument != null) { PdfDocument.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfDocument); }
                if (PdfManager != null) { PdfManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(PdfManager); }
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
