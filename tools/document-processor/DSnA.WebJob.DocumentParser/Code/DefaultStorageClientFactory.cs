//Copyright(c) Microsoft Corporation.All rights reserved.
//Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSnA.WebJob.DocumentParser
{
    public class DefaultStorageClientFactory : IStorageClientFactory
    {
        public const string BlobContainerNameKey = "container";

        public IStorageClient Create(string id, Dictionary<string, string> parameters, IUtils utils)
        {
            switch(id)
            {
                case "blob":
                    return new BlobStorageClient(parameters[BlobContainerNameKey], utils);

                default:
                    return new LocalStorageClient();
            }
        }
    }
}
