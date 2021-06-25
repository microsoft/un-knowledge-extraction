//Copyright(c) Microsoft Corporation.All rights reserved.
//Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSnA.WebJob.DocumentParser
{
    public class StorageObjectDescriptor
    {
        public string FileName { get; set; }
        public Uri Uri { get; set; }
    }

    public interface IStorageClient
    {
        string GetFile(StorageObjectDescriptor descriptor, string destinationFilePath);
        void SaveFile(string sourceUri, string destinationUri);
    }
}
