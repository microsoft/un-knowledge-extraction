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
    public class LocalStorageClient : IStorageClient
    {
        public LocalStorageClient()
        {

        }

        public string GetFile(StorageObjectDescriptor descriptor, string destinationFilePath)
        {
            return System.Net.WebUtility.UrlDecode(descriptor.Uri.AbsolutePath).Replace("/","\\");
        }

        public void SaveFile(string sourceUri, string destinationUri)
        {
            string sourceUriPath = System.IO.Path.GetDirectoryName(sourceUri);
            string sourceUriFileName = System.IO.Path.GetFileName(sourceUri);

            System.IO.File.Copy(sourceUri, (destinationUri != null ? destinationUri : $@"{sourceUriPath}\out_{sourceUriFileName}"), true);
        }
    }
}
