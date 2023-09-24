using KmnlkFileConverterDll.Management;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Web;
using static KmnlkFileConverterApi.Constants.Enums;
using static KmnlkCommon.Shareds.LoggerManagement;

namespace KmnlkFileConverterApi.Management
{
    public class PackageManagement
    {
        private BussinessFileConvertManagement manager;
        public ILog logger;
        public PackageManagement()
        {
            string pathLog = SettingsManagement.getSetting(SettingsManagement.KEY_PathLog).ToString();
            string typeLog = SettingsManagement.getSetting(SettingsManagement.KEY_TypeLog).ToString();
            switch (typeLog.ToLower())
            {
                case "file":
                    logger = new FileLogger(pathLog);
                    break;
                case "db":
                    logger = new DBLogger(pathLog);
                    break;
                default:
                    logger = new FileLogger(pathLog);
                    break;
            }
            manager = new BussinessFileConvertManagement(logger);
        }
        public byte[] convertWordTo(MultipartFormDataStreamProvider provider, int typeExt )
        {
            Guid guid = Guid.NewGuid();
            string dataFolderPath = SettingsManagement.getSetting(SettingsManagement.KEY_DataFolder).ToString();
            dataFolderPath = Path.Combine(dataFolderPath, "UploadWord");

            if (!Directory.Exists(dataFolderPath))
            {
                Directory.CreateDirectory(dataFolderPath);
            }
            dataFolderPath = Path.Combine(dataFolderPath, guid.ToString());
            if (!Directory.Exists(dataFolderPath))
            {
                Directory.CreateDirectory(dataFolderPath);
            }
            byte[] result=null;
            foreach (var file in provider.FileData)
            {
                var name = file.Headers.ContentDisposition.FileName;
                name = name.Trim('"');
                var locationFileName = file.LocalFileName;
                var filePath = Path.Combine(dataFolderPath, guid.ToString()+Path.GetExtension(name));
                File.Copy(locationFileName, filePath);
                string returnPath= manager.convertWordTo(dataFolderPath, filePath, typeExt);
                if(returnPath!=null)
                result = File.ReadAllBytes(returnPath);
                break;
            }
            return result;
        }

        public byte[] convertExcelTo(MultipartFormDataStreamProvider provider, int typeExt)
        {
            Guid guid = Guid.NewGuid();
            string dataFolderPath = SettingsManagement.getSetting(SettingsManagement.KEY_DataFolder).ToString();
            dataFolderPath = Path.Combine(dataFolderPath, "UploadExcel");

            if (!Directory.Exists(dataFolderPath))
            {
                Directory.CreateDirectory(dataFolderPath);
            }
            dataFolderPath = Path.Combine(dataFolderPath, guid.ToString());
            if (!Directory.Exists(dataFolderPath))
            {
                Directory.CreateDirectory(dataFolderPath);
            }
            byte[] result = null;
            foreach (var file in provider.FileData)
            {
                var name = file.Headers.ContentDisposition.FileName;
                name = name.Trim('"');
                var locationFileName = file.LocalFileName;
                var filePath = Path.Combine(dataFolderPath, guid.ToString() + Path.GetExtension(name));
                File.Copy(locationFileName, filePath);
                string returnPath = manager.convertExcelTo(dataFolderPath, filePath, typeExt);
                if (returnPath != null)
                    result = File.ReadAllBytes(returnPath);
                break;
            }
            return result;
        }

        public byte[] convertPdfTo(MultipartFormDataStreamProvider provider, int typeExt)
        {
            Guid guid = Guid.NewGuid();
            string dataFolderPath = SettingsManagement.getSetting(SettingsManagement.KEY_DataFolder).ToString();
            dataFolderPath = Path.Combine(dataFolderPath, "UploadPdf");

            if (!Directory.Exists(dataFolderPath))
            {
                Directory.CreateDirectory(dataFolderPath);
            }
            dataFolderPath = Path.Combine(dataFolderPath, guid.ToString());
            if (!Directory.Exists(dataFolderPath))
            {
                Directory.CreateDirectory(dataFolderPath);
            }
            byte[] result = null;
            foreach (var file in provider.FileData)
            {
                var name = file.Headers.ContentDisposition.FileName;
                name = name.Trim('"');
                var locationFileName = file.LocalFileName;
                var filePath = Path.Combine(dataFolderPath, guid.ToString() + Path.GetExtension(name));
                File.Copy(locationFileName, filePath);
                string returnPath = manager.convertPdfTo(dataFolderPath, filePath, typeExt);
                if (returnPath != null)
                    result = File.ReadAllBytes(returnPath);
                break;
            }
            return result;
        }
        public byte[] convertCompressOrFolderTo(MultipartFormDataStreamProvider provider, int typeExt)
        {
            Guid guid = Guid.NewGuid();
            string dataFolderPath = SettingsManagement.getSetting(SettingsManagement.KEY_DataFolder).ToString();
            dataFolderPath = Path.Combine(dataFolderPath, "UploadFiles");

            if (!Directory.Exists(dataFolderPath))
            {
                Directory.CreateDirectory(dataFolderPath);
            }
            dataFolderPath = Path.Combine(dataFolderPath, guid.ToString());
            if (!Directory.Exists(dataFolderPath))
            {
                Directory.CreateDirectory(dataFolderPath);
            }
            byte[] result = null;
            bool isOk = false;
            foreach (var file in provider.FileData)
            {
                isOk = true;
                var name = file.Headers.ContentDisposition.FileName;
                name = name.Trim('"');
                var locationFileName = file.LocalFileName;
                var filePath = Path.Combine(dataFolderPath,name);
                File.Copy(locationFileName, filePath);
            }
            if (isOk)
            {
                string returnPath = manager.convertCompressOrFolderTo(dataFolderPath, dataFolderPath, typeExt);
                if (returnPath != null)
                    result = File.ReadAllBytes(returnPath);
            }
            return result;
        }
    }
}