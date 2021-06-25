//Copyright(c) Microsoft Corporation.All rights reserved.
//Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DSnA.WebJob.DocumentParser
{
    public interface IStorageClientFactory
    {
        IStorageClient Create(string id, Dictionary<string, string> parameters, IUtils utils);
    }
}
