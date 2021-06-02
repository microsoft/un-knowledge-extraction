//Copyright(c) Microsoft Corporation.All rights reserved.
//Licensed under the MIT License.

using System;

namespace DSnA.WebJob.DocumentParser
{
    using System.Diagnostics;
    using Microsoft.Azure;
    using Microsoft.WindowsAzure.Storage;
    using Microsoft.WindowsAzure.Storage.Blob;

    public interface ILogger
    {
        void Info(string message);
        void Error(string message, Exception exp);
    }
    public class Logger : ILogger
    {
        private static CloudStorageAccount StorageAccount = CloudStorageAccount.Parse(CloudConfigurationManager.GetSetting("StorageConnectionString"));

        private static CloudBlobClient blobClient = StorageAccount.CreateCloudBlobClient();

        private static string logContainerName = CloudConfigurationManager.GetSetting(Constants.ParserConfig.LogsContainerNameRef);

        private static string logPrefix = CloudConfigurationManager.GetSetting(Constants.ParserConfig.LogPrefix);

        private static Logger LoggerInstance;
        private Logger() { }

        public static Logger Instance
        {
            get
            {
                if (LoggerInstance == null)
                {
                    LoggerInstance = new Logger();
                }

                return LoggerInstance;
            }
        }

        /// <summary>
        /// Write log text (info/error) to Azure Blob
        /// </summary>
        /// <param name="message"></param>
        /// <param name="category"></param>
        /// <param name="exp"></param>
        private void Write(string message, EventLogEntryType category, Exception exp = null)
        {
            try
            {
                // create blob client and container(if not exists) to store the logs in Azure Storage Account
                CloudBlobContainer logsContainer = blobClient.GetContainerReference(logContainerName);
                logsContainer.CreateIfNotExists();
                // append information to blob - create logs for every day if not exists
                CloudAppendBlob appendBlob = logsContainer.GetAppendBlobReference($"log_{logPrefix}_{DateTime.UtcNow.ToString("MM-dd-yyyy")}.log");
                if (!appendBlob.Exists())
                    appendBlob.CreateOrReplace();

                if (exp != null)
                    appendBlob.AppendText(String.Format("{0:u}\t[{1}]\t[{2}]\tMessage:{3}{4}{5}{6}",
                        DateTime.UtcNow, Environment.MachineName, category.ToString().ToUpper(), message, Environment.NewLine, exp, Environment.NewLine));
                else
                    appendBlob.AppendText(String.Format("{0:u}\t[{1}]\t[{2}]\tMessage:{3}{4}",
                        DateTime.UtcNow, Environment.MachineName, category.ToString().ToUpper(), message, Environment.NewLine));
            }
            catch (Exception exception)
            {
                throw new LoggerException("Exception in Logging information/error", exception);
            }
        }

        /// <summary>
        /// Logs information text
        /// </summary>
        /// <param name="message"></param>
        public void Info(string message)
        {
            Write(message, EventLogEntryType.Information);
        }

        /// <summary>
        /// Logs Exception text
        /// </summary>
        /// <param name="message"></param>
        /// <param name="exp"></param>
        public void Error(string message, Exception exp)
        {
            Write(message, EventLogEntryType.Error, exp);
        }
    }
}
