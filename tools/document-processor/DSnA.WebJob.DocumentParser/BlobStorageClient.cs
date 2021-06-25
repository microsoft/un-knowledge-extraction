//Copyright(c) Microsoft Corporation.All rights reserved.
//Licensed under the MIT License.

using Microsoft.Azure;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System;
using System.Diagnostics;
using System.Linq;

namespace DSnA.WebJob.DocumentParser
{
    public class BlobStorageClient : IStorageClient
    {
        private static readonly CloudStorageAccount StorageAccount = CloudStorageAccount.Parse(CloudConfigurationManager.GetSetting("StorageConnectionString"));

        private readonly IUtils _utils;

        public CloudBlobClient Client;
        public CloudBlobContainer Container;

        public BlobStorageClient(string containerName, IUtils utils)
        {
            _utils = utils;

            Client = utils.CreateCloudBlobClient(StorageAccount);
            Container = Client.GetContainerReference(containerName);
        }

        public string GetFile(StorageObjectDescriptor descriptor, string destinationFilePath)
        {
            return _utils.DownloadBlobFile(descriptor.Uri.AbsoluteUri, Constants.FileConfigs.WorkingDirectoryPath, Client);
        }

        public void SaveFile(string sourceUri, string destinationUri)
        {
            _utils.UploadFileToBlob(sourceUri, Client);
        }
    }
}
