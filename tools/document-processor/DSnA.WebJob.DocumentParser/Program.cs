//Copyright(c) Microsoft Corporation.All rights reserved.
//Licensed under the MIT License.

using Microsoft.Azure;
using Microsoft.Office.Interop.Word;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Blob;
using System;
using System.Diagnostics;
using System.Linq;

namespace DSnA.WebJob.DocumentParser
{
    class Program
    {
        private static string _storageType = CloudConfigurationManager.GetSetting("StorageType");

        static void Main(string[] args)
        {
            if (!ValidateArgs(args, _storageType))
                return;

            var logger = ConsoleLogger.Instance;
            var util = new Utils(logger);
            
            var args0 = _storageType == "blob" ? args[0] : Constants.FileConfigs.SourceDirectoryPath = args[0];
            var args1 = _storageType == "blob" ? args[1] : Constants.FileConfigs.OutputDirectoryPath = args[1];
            var args2 = args.Count() > 2 ? args[2] : null;
            
            IStorageClientFactory clientFactory = new DefaultStorageClientFactory();
            IStorageClient client = clientFactory.Create(_storageType, new System.Collections.Generic.Dictionary<string, string>() {
                                        { DefaultStorageClientFactory.BlobContainerNameKey, args0 }
                                    }, util);

            string[] uris = GetUris(client, _storageType == "blob" ? args1 : Constants.FileConfigs.SourceDirectoryPath, args2, util);

            IDocumentParser parser = new DocumentParser(logger, util, new ParseHelper(logger, util));

            var total = uris.Count();
            if (total == 0)
            {
                Console.WriteLine("No files to process...");
            }

            var counter = 0;

            var outputFileFormat = CloudConfigurationManager.GetSetting("OutputFileFormat");

            Stopwatch stopWatch = new Stopwatch();

            InteropWordUtils iInteropWordUtils = new InteropWordUtils();

            // fire up word instance
            Application wordApp = iInteropWordUtils.CreateWordAppInstance();

            int maxFailures = 3;
            int currentFailures = 0;

            try
            {
                foreach (var uri in uris)
                {
                    stopWatch.Start();
                    counter++;
                    Console.WriteLine($"Processing: {counter} out of {total}");
                    Console.WriteLine($"Processing: {uri}");

                    string result = null;

                    try
                    {
                        result = parser.ParseDocuments(uri, client, wordApp, outputFileFormat);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        Console.WriteLine(ex.StackTrace);

                        currentFailures++;

                        if (currentFailures >= maxFailures)
                        {
                            throw new Exception("Max failure count reached.");
                        }
                    }

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
            finally
            {
                iInteropWordUtils.DisposeIneropObject(wordApp);
            }

            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }

        private static string[] GetUris(IStorageClient client, string prefix, string filter, IUtils utils)
        {
            if (client is BlobStorageClient) return GetBlobUris(client as BlobStorageClient, prefix, filter, utils);
            else if (client is LocalStorageClient) return GetLocalUris(client as LocalStorageClient, prefix, filter, utils);
            else return null;
        }

        private static string[] GetBlobUris(BlobStorageClient client, string sourcePath, string filter, IUtils utils)
        {
            Console.WriteLine($"Listing Blobs in container {client.Container.Name} in folder {sourcePath}");

            string blobPrefix = sourcePath == "null" || sourcePath == "/" ? null : sourcePath;

            var outputBlobList = utils.GetBlobListFromOutputContainer(client.Client);

            var blobList = client.Container.ListBlobs(prefix: blobPrefix, useFlatBlobListing: true);

            var filteredBlobList = blobList.Where(s => !outputBlobList.Contains(utils.CleanNonSupportedSparkChar(s.Uri.Segments[s.Uri.Segments.Length - 1]))).ToList();

            if (filter != null)
            {
                filteredBlobList = filteredBlobList.Where(s => s.Uri.PathAndQuery.Contains(filter)).ToList();
            }

            return filteredBlobList
                .Select(x => x.Uri.AbsoluteUri)
                .ToArray();
        }

        private static string[] GetLocalUris(LocalStorageClient client, string sourcePath, string filter, IUtils utils)
        {
            string[] files = System.IO.Directory.GetFiles(sourcePath, "*", System.IO.SearchOption.AllDirectories);

            return (!string.IsNullOrEmpty(filter)
                ? files.Where(x => x.Contains(filter)).ToArray()
                : files);
        }

        private static bool ValidateArgs(string[] args, string storageType)
        {
            bool validArgs = true;

            switch (storageType)
            {
                case "blob":
                    if (!args?.Any() ?? true)
                    {
                        Console.WriteLine(string.Format(" No arguments passed. \n\n DSnA.WebJob.DocumentParser.exe arg1 arg2 arg3 \n\n Options: \n\t arg1: Required - blob container name \n\t arg2: Required - virtual directory name (/ root level) \n\t arg3: Optional - document file name filter"));
                        validArgs = false;
                    }
                    else if (!(args.Length >= 2 && args.Length < 4))
                    {
                        Console.WriteLine(string.Format(" Incorrect number of arguments. \n\n DSnA.WebJob.DocumentParser.exe arg1 arg2 arg3 \n\n Options: \n\t arg1: Required - blob container name \n\t arg2: Required - virtual directory name (/ root level) \n\t arg3: Optional - document file name filter"));
                        validArgs = false;
                    }
                    break;

                case "localstorage":
                    if (!args?.Any() ?? true)
                    {
                        Console.WriteLine(string.Format(" No arguments passed. \n\n DSnA.WebJob.DocumentParser.exe arg1 arg2 arg3 \n\n Options:  \n\t arg1: Required - local storage source folder path \n\t arg2: Required - local storage output folder path \n\t arg3: Optional - document file name filter"));
                        validArgs = false;
                    }
                    else if (!(args.Length >= 2 && args.Length < 4))
                    {
                        Console.WriteLine(string.Format(" Incorrect number of arguments. \n\n DSnA.WebJob.DocumentParser.exe arg1 arg2 arg3 \n\n Options:  \n\t arg1: Required - local storage source folder path \n\t arg2: Required - local storage output folder path \n\t arg3: Optional - document file name filter"));
                        validArgs = false;
                    }
                    break;

                default:
                    return validArgs = false;
            }

            return validArgs;
        }
    }
}
