
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

namespace KmnlkFileConverterDll.Management
{
    public class WordConvertManagement : IWordConvertOperations, IValidationOperations
    {
        private ILog logger;
        private Microsoft.Office.Interop.Word.Application wordManager;
        private Microsoft.Office.Interop.Word.Document wordDocument;
        public WordConvertManagement(ILog logger)
        {
            this.logger = logger;
            wordManager = new Microsoft.Office.Interop.Word.Application();
        }

        public string convertWordToExcel(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                wordManager = new Microsoft.Office.Interop.Word.Application();
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".xlsx";
                wordDocument = wordManager.Documents.Open(path);
              //  wordDocument.SaveAs2(newPath, WdSaveFormat.wdFormatDocument);
                wordDocument.Close(false, false, false);
                wordManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return "";
            }
            catch (Exception e)
            {
                if (wordDocument != null) { wordDocument.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument); }
                if (wordManager != null) { wordManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(wordManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string convertWordToHtml(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                wordManager = new Microsoft.Office.Interop.Word.Application();
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".html";
                wordDocument = wordManager.Documents.Open(path);
                wordDocument.SaveAs2(newPath, WdSaveFormat.wdFormatFilteredHTML);
                wordDocument.Close(false, false, false);
                wordManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (wordDocument != null) { wordDocument.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument); }
                if (wordManager != null) { wordManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(wordManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string convertWordToXml(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                wordManager = new Microsoft.Office.Interop.Word.Application();
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".xml";
                wordDocument = wordManager.Documents.Open(path);
                wordDocument.SaveAs2(newPath, WdSaveFormat.wdFormatXML);
                wordDocument.Close(false, false, false);
                wordManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (wordDocument != null) { wordDocument.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument); }
                if (wordManager != null) { wordManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(wordManager); } 
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }
        public string convertWordToRtf(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                wordManager = new Microsoft.Office.Interop.Word.Application();
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".rtf";
                wordDocument = wordManager.Documents.Open(path);
                wordDocument.SaveAs2(newPath, WdSaveFormat.wdFormatRTF);
                wordDocument.Close(false, false, false);
                wordManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (wordDocument != null) { wordDocument.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument); }
                if (wordManager != null) { wordManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(wordManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }
        public string convertWordToPdf(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                wordManager = new Microsoft.Office.Interop.Word.Application();
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath+ ".pdf";
                 wordDocument = wordManager.Documents.Open(path);
                //wordDocument.ExportAsFixedFormat(newPath, WdExportFormat.wdExportFormatPDF);
                wordDocument.SaveAs2(newPath, WdSaveFormat.wdFormatPDF);
                wordDocument.Close(false,false,false);
                wordManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (wordDocument != null) { wordDocument.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument); }
                if (wordManager != null) { wordManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(wordManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string convertWordToText(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                wordManager = new Microsoft.Office.Interop.Word.Application();
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".txt";
                wordDocument = wordManager.Documents.Open(path);
                wordDocument.SaveAs2(newPath, WdSaveFormat.wdFormatText);
                wordDocument.Close(false, false, false);
                wordManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (wordDocument != null) { wordDocument.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument); }
                if (wordManager != null) { wordManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(wordManager); }
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string convertWordToWord(string dataFolderPath, string path)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), path, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(path))
                {
                    return null;
                }
                wordManager = new Microsoft.Office.Interop.Word.Application();
                Guid guid = Guid.NewGuid();
                string nameFile = guid.ToString();
                string tempPath = MainHelper.getPathWithOutExt(path);
                string newPath = tempPath + ".docx";
                wordDocument = wordManager.Documents.Open(path);
                wordDocument.SaveAs2(newPath, WdSaveFormat.wdFormatText);
                wordDocument.Close(false, false, false);
                wordManager.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordManager);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                if (wordDocument != null) { wordDocument.Close(); System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument); }
                if (wordManager != null) { wordManager.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(wordManager); }
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
