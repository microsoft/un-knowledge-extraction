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
    class Program
    {
        private static readonly CloudStorageAccount StorageAccount = CloudStorageAccount.Parse(CloudConfigurationManager.GetSetting("StorageConnectionString"));
        static void Main(string[] args)
        {
            if (!ValidateArgs(args))
                return;

            var logger = Logger.Instance;
            var util = new Utils(Logger.Instance);

            CloudBlobClient blobClient = util.CreateCloudBlobClient(StorageAccount);

            CloudBlobContainer container = blobClient.GetContainerReference(args[0]);

            Console.WriteLine($"Listing Blobs in container {args[0]} in folder {args[1]}");

            string blobPrefix = args[1] == "/" ? null : args[1];

            var outputBlobList = util.GetBlobListFromOutputContainer(blobClient);

            var blobList = container.ListBlobs(prefix: blobPrefix, useFlatBlobListing: true);

            var filteredBlobList = blobList.Where(s => !outputBlobList.Contains(util.CleanNonSupportedSparkChar(s.Uri.Segments[s.Uri.Segments.Length - 1]))).ToList();

            if (args.Length == 3 && args[2] != null)
            {
                filteredBlobList = filteredBlobList.Where(s => s.Uri.PathAndQuery.Contains(args[2])).ToList();
            }

            IDocumentParser parser = new DocumentParser(logger, util, new ParseHelper(logger, util));

            var total = filteredBlobList.Count();

            var counter = 0;

            var outputFileFormat = CloudConfigurationManager.GetSetting("OutputFileFormat");

            Stopwatch stopWatch = new Stopwatch();

            foreach (var item in filteredBlobList)
            {
                stopWatch.Start();
                counter++;
                Console.WriteLine($"Processing: {counter} out of {total}");
                Console.WriteLine($"Processing: {item.Uri.ToString()}");

                var result = parser.ParseDocuments(item.Uri.ToString(), outputFileFormat, blobClient);

                stopWatch.Stop();

                TimeSpan ts = stopWatch.Elapsed;

                string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                                                    ts.Hours, ts.Minutes, ts.Seconds,
                                                    ts.Milliseconds / 10);

                Console.WriteLine("RunTime " + elapsedTime);
                Console.WriteLine(result);
                stopWatch.Reset();
            }
        }

        private static bool ValidateArgs(string[] args)
        {
            bool validArgs = true;

            if (!args?.Any() ?? true)
            {
                Console.WriteLine(string.Format(" No arguments passed. \n\n DSnA.WebJob.DocumentParser.exe arg1 arg2 arg3 \n\n Options: \n\t arg1: Required - blob container name \n\t arg2: Required - virtual directory name (/ root level) \n\t arg3: Optional - file name filter"));
                validArgs = false;
            }
            else if (!(args.Length >= 2 && args.Length < 4))
            {
                Console.WriteLine(string.Format(" Incorrect number of arguments. \n\n DSnA.WebJob.DocumentParser.exe arg1 arg2 arg3 \n\n Options: \n\t arg1: Required - blob container name \n\t arg2: Required - virtual directory name (/ root level) \n\t arg3: Optional - file name filter"));
                validArgs = false;
            }

            return validArgs;
        }
    }
}
