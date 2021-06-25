//Copyright(c) Microsoft Corporation.All rights reserved.
//Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace DSnA.WebJob.DocumentParser
{
    public interface IDocumentParser
    {
        string ParseDocuments(string uri, IStorageClient storageClient, Application wordApp, string outputFileFormat);
    }
}
