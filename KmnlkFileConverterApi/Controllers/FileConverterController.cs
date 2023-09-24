using KmnlkFileConverterApi.Constants;
using KmnlkFileConverterApi.Exceptions;
using KmnlkFileConverterApi.Management;
using KmnlkFileConverterApi.Models;
using KmnlkCommon.Shareds;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;
using static KmnlkCommon.Shareds.LoggerManagement;
using KmnlkFileConverterDll.Helpers;

namespace KmnlkFileConverterApi.Controllers
{
    public class FileConverterController : ApiController
    {
        private PackageManagement package = null;

        public FileConverterController(PackageManagement repo)
        {
            package = repo;
        }

        [HttpPost]
        [ActionName("ConvertWordTo")]
        public async Task<HttpResponseMessage> ConvertWordTo([FromUri]int type, string fileName = "download")
        {
            package.logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstants.MSG_SUCCESS);
            string startTime = DateTime.Now.ToString("hh:mm:ss");
            string endTime = "";
            HttpResponseMessage res;
            try
            {
                string dataFolder = SettingsManagement.getSetting(SettingsManagement.KEY_DataFolder).ToString();
                dataFolder = Path.Combine(dataFolder, "TempFolderWord");
                if (!Directory.Exists(dataFolder))
                {
                    Directory.CreateDirectory(dataFolder);
                }
                var provider = new MultipartFormDataStreamProvider(dataFolder);
                await Request.Content.ReadAsMultipartAsync(provider);
                byte[] bytesFile = package.convertWordTo(provider, type);
                endTime = DateTime.Now.ToString("hh:mm:ss");
                res = DownloadManagement.Download(bytesFile, fileName + "." + MainHelper.getStringTypeExt(type), MainHelper.getStringTypeExt(type), MainHelper.getStringTypeExt(type));
                package.logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstants.MSG_SUCCESS);
                return res;
            }
            catch (Exception e)
            {
                new ApiException(package.logger, modConstants.MSG_SUCCESS, EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                var response = new ResponseModel(null, e.Message, HttpStatusCode.BadRequest, startTime, endTime);
                return Request.CreateResponse<ResponseModel>(HttpStatusCode.OK, response);
            }

        }

        [HttpPost]
        [ActionName("ConvertExcelTo")]
        public async Task<HttpResponseMessage> ConvertExcelTo([FromUri]int type, string fileName = "download")
        {
            package.logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstants.MSG_SUCCESS);
            string startTime = DateTime.Now.ToString("hh:mm:ss");
            string endTime = "";
            HttpResponseMessage res;
            try
            {
                string dataFolder = SettingsManagement.getSetting(SettingsManagement.KEY_DataFolder).ToString();
                dataFolder = Path.Combine(dataFolder, "TempFolderExcel");
                if (!Directory.Exists(dataFolder))
                {
                    Directory.CreateDirectory(dataFolder);
                }
                var provider = new MultipartFormDataStreamProvider(dataFolder);
                await Request.Content.ReadAsMultipartAsync(provider);
                byte[] bytesFile = package.convertExcelTo(provider, type);
                endTime = DateTime.Now.ToString("hh:mm:ss");
                res = DownloadManagement.Download(bytesFile, fileName + "." + MainHelper.getStringTypeExt(type), MainHelper.getStringTypeExt(type), MainHelper.getStringTypeExt(type));
                package.logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstants.MSG_SUCCESS);
                return res;
            }
            catch (Exception e)
            {
                new ApiException(package.logger, modConstants.MSG_SUCCESS, EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                var response = new ResponseModel(null, e.Message, HttpStatusCode.BadRequest, startTime, endTime);
                return Request.CreateResponse<ResponseModel>(HttpStatusCode.OK, response);
            }

        }

        [HttpPost]
        [ActionName("ConvertPdfTo")]
        public async Task<HttpResponseMessage> ConvertPdfTo([FromUri]int type, string fileName = "download")
        {
            package.logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstants.MSG_SUCCESS);
            string startTime = DateTime.Now.ToString("hh:mm:ss");
            string endTime = "";
            HttpResponseMessage res;
            try
            {
                string dataFolder = SettingsManagement.getSetting(SettingsManagement.KEY_DataFolder).ToString();
                dataFolder = Path.Combine(dataFolder, "TempFolderPdf");
                if (!Directory.Exists(dataFolder))
                {
                    Directory.CreateDirectory(dataFolder);
                }
                var provider = new MultipartFormDataStreamProvider(dataFolder);
                await Request.Content.ReadAsMultipartAsync(provider);
                byte[] bytesFile = package.convertPdfTo(provider, type);
                endTime = DateTime.Now.ToString("hh:mm:ss");
                res = DownloadManagement.Download(bytesFile, fileName + "." + MainHelper.getStringTypeExt(type), MainHelper.getStringTypeExt(type), MainHelper.getStringTypeExt(type));
                package.logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstants.MSG_SUCCESS);
                return res;
            }
            catch (Exception e)
            {
                new ApiException(package.logger, modConstants.MSG_SUCCESS, EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                var response = new ResponseModel(null, e.Message, HttpStatusCode.BadRequest, startTime, endTime);
                return Request.CreateResponse<ResponseModel>(HttpStatusCode.OK, response);
            }

        }

        [HttpPost]
        [ActionName("Compress")]
        public async Task<HttpResponseMessage> Compress([FromUri]int type, string fileName = "download")
        {
            package.logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstants.MSG_SUCCESS);
            string startTime = DateTime.Now.ToString("hh:mm:ss");
            string endTime = "";
            HttpResponseMessage res;
            try
            {
                string dataFolder = SettingsManagement.getSetting(SettingsManagement.KEY_DataFolder).ToString();
                dataFolder = Path.Combine(dataFolder, "TempFolderCompress");
                if (!Directory.Exists(dataFolder))
                {
                    Directory.CreateDirectory(dataFolder);
                }
                var provider = new MultipartFormDataStreamProvider(dataFolder);
                await Request.Content.ReadAsMultipartAsync(provider);
                byte[] bytesFile = package.convertCompressOrFolderTo(provider, type);
                endTime = DateTime.Now.ToString("hh:mm:ss");
                res = DownloadManagement.Download(bytesFile, fileName + "." + MainHelper.getStringTypeExt(type), MainHelper.getStringTypeExt(type), MainHelper.getStringTypeExt(type));
                package.logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstants.MSG_SUCCESS);
                return res;
            }
            catch (Exception e)
            {
                new ApiException(package.logger, modConstants.MSG_SUCCESS, EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                var response = new ResponseModel(null, e.Message, HttpStatusCode.BadRequest, startTime, endTime);
                return Request.CreateResponse<ResponseModel>(HttpStatusCode.OK, response);
            }

        }
        [NonAction]
        public bool isValid(string uid)
        {
            return true;
        }
    }
}
