
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
using System.IO.Compression;
namespace KmnlkFileConverterDll.Management
{
    public class CompressConvertManagement : ICompressConvertOperations, IValidationOperations
    {
        private ILog logger;

        public CompressConvertManagement(ILog logger)
        {
            this.logger = logger;
        }

        public string convertFolderToZipFile(string dataFolderPath, string pathSource)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), pathSource, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(pathSource))
                {
                    return null;
                }
                Guid guid = Guid.NewGuid();
                string newPath = pathSource + ".zip";
                ZipFile.CreateFromDirectory(pathSource, newPath);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }


        public string extractFolderFromZipFile(string dataFolderPath, string pathZip)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), pathZip, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(pathZip))
                {
                    return null;
                }
                Guid guid = Guid.NewGuid();
                string newPath = MainHelper.getPathWithOutExt(pathZip);
                ZipFile.ExtractToDirectory(pathZip, newPath);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string convertFolderToRarFile(string dataFolderPath, string pathSource)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), pathSource, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(pathSource))
                {
                    return null;
                }
                Guid guid = Guid.NewGuid();
                string newPath = pathSource + ".rar";
                ZipFile.CreateFromDirectory(pathSource, newPath);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
                new DllException(logger, "", EnvironmentManagement.getCurrentMethodName(this.GetType()), e.Message);
                return null;
            }
        }

        public string extractFolderFromRarFile(string dataFolderPath, string pathZip)
        {
            logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), pathZip, ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.START, modConstant.MSG_SUCCESS);
            try
            {
                if (!isValid(pathZip))
                {
                    return null;
                }
                Guid guid = Guid.NewGuid();
                string newPath = MainHelper.getPathWithOutExt(pathZip);
                ZipFile.ExtractToDirectory(pathZip, newPath);
                logger.WriteToLog(EnvironmentManagement.getCurrentMethodName(this.GetType()), "", ENUM_TYPE_MSG_LOGGER.INFO, ENUM_TYPE_Block_LOGGER.END, modConstant.MSG_SUCCESS);
                return newPath;
            }
            catch (Exception e)
            {
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
