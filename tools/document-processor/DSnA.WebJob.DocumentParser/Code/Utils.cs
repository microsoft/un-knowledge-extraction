//Copyright(c) Microsoft Corporation.All rights reserved.
//Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Word;
using Newtonsoft.Json;
using System.Text.RegularExpressions;
using System.Linq;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.WindowsAzure.Storage.Queue;


namespace DSnA.WebJob.DocumentParser
{
    using Microsoft.Azure;
    using Microsoft.WindowsAzure.Storage;
    using System.IO.Abstractions;
    using Table = Table;
    public interface IUtils
    {
        string CleanNonSupportedSparkChar(string dirtyString);
        string CleanTextFromNonAsciiChar(string dirtyString);
        List<string> ExtractLinksFromText(string content, bool isPreProcessingReq = false);
        string SerializeAndSaveJson(dynamic jsonData, string fileName);
        string SaveToCsvFile(List<string> csvLines, string fileName);
        string SaveJsonToFile(string jsonDoc, string fileName, string directory);
        FileMetaData ExtractFileMetadata(string fileLocation);
        void UploadFileToBlob(string fileLocation, CloudBlobClient blobClient);
        List<string> GetBlobListFromOutputContainer(CloudBlobClient blobClient);
        void CheckBlobUriInMsg(QueueMessage queueMsg);
        void DeleteInputFiles(List<string> filesToDelete);
        void CheckAllQueueExists(List<CloudQueue> queueList);
        string DownloadBlobFile(string blobUri, string locationToSave, CloudBlobClient blobClient);
        bool CheckUriIsValid(string inputUri);
        string ConvertPdfToWord(string file, string directoryToSave);
        JsonDocumentStruct PrepareErrorJsonDoc(string fileLocation, Exception exp);
        Tuple<int, int> FindTableWithHeader(Document wordDocToExtract, List<string> tableHeaders);
        CloudBlobClient CreateCloudBlobClient(CloudStorageAccount StorageAccount);
        CloudQueue GetQueueReference(CloudStorageAccount StorageAccount);
        int GetQueueMessageDequeueCount(CloudQueueMessage queueMsg);
    }

    public class Utils : IUtils
    {
        private readonly ILogger iLogger;
        private readonly IFileSystem iFileSystem;
        private readonly IInteropWordUtils iInteropWordUtils;
        private static string outputContainerName = CloudConfigurationManager.GetSetting(Constants.ParserConfig.OutputContainerNameRef);

        public Utils(ILogger iLogger)
        {
            this.iLogger = iLogger;
            this.iFileSystem = new FileSystem();
            iInteropWordUtils = new InteropWordUtils();
        }

        public Utils(ILogger iLogger, IFileSystem iFileSystem)
        {
            this.iLogger = iLogger;
            this.iFileSystem = iFileSystem;
            iInteropWordUtils = new InteropWordUtils();
        }

        public Utils(ILogger iLogger, IFileSystem iFileSystem, IInteropWordUtils iInteropWordUtils)
        {
            this.iLogger = iLogger;
            this.iFileSystem = iFileSystem;
            this.iInteropWordUtils = iInteropWordUtils;
        }

        /// <summary>
        /// Clean non supported Spark filename characters
        /// </summary>
        /// <param name="dirtyString"></param>
        /// <returns>text with only ascii char</returns>
        public string CleanNonSupportedSparkChar(string dirtyString)
        {
            if (string.IsNullOrEmpty(dirtyString))
                return dirtyString;

            return dirtyString.Replace("%20", "_").Replace("{", "").Replace("}", "").Replace("[", "").Replace("]", "");
        }

        /// <summary>
        /// Clean non ascii char from input text/string
        /// </summary>
        /// <param name="dirtyString"></param>
        /// <returns>text with only ascii char</returns>
        public string CleanTextFromNonAsciiChar(string dirtyString)
        {
            if (string.IsNullOrEmpty(dirtyString))
                return dirtyString;

            string cleanString = Regex.Replace(dirtyString, Constants.RegexExp.NoEscapeSequences, String.Empty);
            cleanString = Regex.Replace(cleanString, Constants.RegexExp.OnlyAsciiChar, String.Empty);
            cleanString = Regex.Replace(cleanString, "\u0001", String.Empty);
            cleanString = Regex.Replace(cleanString, "\u0015", String.Empty);
            cleanString = Regex.Replace(cleanString, Constants.RegexExp.OnlyWhiteSpaces, " ");
            return cleanString.Trim();
        }

        /// <summary>
        /// extract only hyperlinks from text
        /// if preprocessing = true -> removes all spaces and adds spaces only before App protocols
        /// -to distinguish hyperlinks from other strings.
        /// </summary>
        /// <param name="content"></param>
        /// <param name="isPreProcessingReq"></param>
        /// <returns>list of hyperlinks</returns>
        public List<string> ExtractLinksFromText(string content, bool isPreProcessingReq = false)
        {
            // do some preprocessing on the text
            if (isPreProcessingReq)
            {
                string stringWithSpacesRemoved = Regex.Replace(content, Constants.RegexExp.OnlyWhiteSpaces, String.Empty);
                content = Regex.Replace(stringWithSpacesRemoved, Constants.RegexExp.HyperlinkAppProtocols, word => String.Format(@" {0}", word.Value)); //add space before http or https or ftp so that next regex can pick the link
            }

            MatchCollection matches = Regex.Matches(content, Constants.RegexExp.OnlyHyperlinks);
            List<string> webLinks = matches.Cast<Match>().Select(match => match.Value).ToList();
            return webLinks;
        }

        /// <summary>
        /// Serialize json string and save it to location
        /// </summary>
        /// <param name="jsonData"></param>
        /// <param name="fileName"></param>
        /// <returns>location where json file is saved</returns>
        public string SerializeAndSaveJson(dynamic jsonData, string fileName)
        {
            try
            {
                if (String.IsNullOrEmpty(fileName))
                    fileName = Constants.FileConfigs.TempFileName + "-" + DateTime.UtcNow.ToString(Constants.DateTimeFormat);

                string finalJson = JsonConvert.SerializeObject(jsonData, Formatting.Indented);
                var jsonOutputFileLocation = SaveJsonToFile(finalJson, fileName, Constants.FileConfigs.OutputDirectoryPath);
                return jsonOutputFileLocation;
            }
            catch (Exception exception)
            {
                throw new Exception("Exception occured in saving JSON (SerializeAndSaveJson)\n", exception);
            }
        }

        public string SaveToCsvFile(List<string> csvLines, string fileName)
        {
            if (String.IsNullOrEmpty(fileName))
                fileName = Constants.FileConfigs.TempFileName + "-" + DateTime.UtcNow.ToString(Constants.DateTimeFormat);

            var directory = Constants.FileConfigs.OutputDirectoryPath;
            if (!iFileSystem.Directory.Exists(directory))
                iFileSystem.Directory.CreateDirectory(directory);

            string outputFileName = Path.Combine(directory, $"{fileName}.csv");
            using (StreamWriter sw = new StreamWriter(outputFileName))
            {
                foreach (var row in csvLines)
                {
                    sw.WriteLine(row);
                }
            }

            return outputFileName;
        }

        /// <summary>
        /// Populate all metadata related to extracted file
        /// </summary>
        /// <param name="fileLocation"></param>
        /// <param name="queueMessage"></param>
        /// <returns>FileMetaData</returns>
        public FileMetaData ExtractFileMetadata(string fileLocation)
        {
            try
            {
                FileMetaData fileData = new FileMetaData();
                var filName = Path.GetFileName(fileLocation);
                fileData.AgreementNumber = filName.Contains("_") ? filName.Split('_')[0] : string.Empty;
                fileData.FileName = filName;
                fileData.FileType = Path.GetExtension(fileLocation).Replace(".","");
                fileData.ExtractionTimeStamp = DateTime.UtcNow.ToString(Constants.DateTimeFormat);
                return fileData;
            }
            catch (Exception exception)
            {
                throw new Exception("Error in extracting metadata of given file(ExtractFileMetadata)\n", exception);
            }
        }

        /// <summary>
        /// upload file to given blob location
        /// </summary>
        /// <param name="fileLocation"></param>
        /// <param name="blobClient"></param>
        public void UploadFileToBlob(string fileLocation, CloudBlobClient blobClient)
        {
            try
            {
                var container = blobClient.GetContainerReference(outputContainerName);
                ICloudBlob blob = container.GetBlockBlobReference(Path.GetFileName(fileLocation));
                if (blob == null)
                    throw new Exception("Inaccessible blob location --> " + container?.Uri?.AbsolutePath + " (UploadFileToBlob)\n");

                blob.UploadFromFile(fileLocation);
            }
            catch (Exception exception)
            {
                throw new Exception("Error in uploading the output JSON file to blob location (" + fileLocation + ") in (UploadFileToBlob)\n", exception);
            }
        }

        public List<string> GetBlobListFromOutputContainer(CloudBlobClient blobClient)
        {
            try
            {
                var container = blobClient.GetContainerReference(outputContainerName);
                var blobList = container.ListBlobs(useFlatBlobListing: true);
                var ouputBlobs = blobList.Select(s => s.Uri.Segments[s.Uri.Segments.Length - 1].Replace(".json","")).ToList();
                return ouputBlobs;
            }
            catch (Exception exception)
            {
                throw new Exception("Error while executing func GetBlobListFromOutputContainer", exception);
            }
        }

        /// <summary>
        /// higher level - check blob URI is valid
        /// </summary>
        /// <param name="queueMsg"></param>
        public void CheckBlobUriInMsg(QueueMessage queueMsg)
        {
            if (!CheckUriIsValid(queueMsg.FileInputUri))
                throw new Exception("Queue message is invalid" + "-->" + queueMsg.FileInputUri);

            if (!CheckUriIsValid(queueMsg.FileOutputUri))
                throw new Exception("Queue message is invalid" + "-->" + queueMsg.FileOutputUri);
        }

        /// <summary>
        /// Delete given list of files
        /// </summary>
        /// <param name="filesToDelete"></param>
        public void DeleteInputFiles(List<string> filesToDelete)
        {
            try
            {
                foreach (var file in filesToDelete)
                {
                    if (iFileSystem.File.Exists(file))
                        iFileSystem.File.Delete(file);
                }
            }
            catch (Exception exception)
            {
                throw new UnableToDeleteFileException("Unable to delete files related to reports (pdf doc or word doc or json output file)\n", exception);
            }
        }

        /// <summary>
        /// Check given list of Azure Queues exist
        /// </summary>
        /// <param name="queueList"></param>
        public void CheckAllQueueExists(List<CloudQueue> queueList)
        {
            foreach (CloudQueue queue in queueList)
            {
                if (!queue.Exists())
                    throw new Exception("Message Queue is inaccessible or does not exist" + "-->" + queue.Uri + "(CheckAllQueueExists)\n");
            }
        }

        /// <summary>
        /// Download blob file to local file system 
        /// </summary>
        /// <param name="blobUri"></param>
        /// <param name="locationToSave"></param>
        /// <param name="blobClient"></param>
        /// <returns>local file location where the file is saved</returns>
        public string DownloadBlobFile(string blobUri, string locationToSave, CloudBlobClient blobClient)
        {
            try
            {
                if (!iFileSystem.Directory.Exists(locationToSave))
                    iFileSystem.Directory.CreateDirectory(locationToSave);
                
                ICloudBlob blob = blobClient.GetBlobReferenceFromServer(new Uri(blobUri));
                if (blob == null)
                    throw new Exception("Inaccessible blob location " + blobUri + "\n");

                string fileName = Path.GetFileName(new Uri(blobUri).LocalPath).Replace(" ", "_").Replace("{", "").Replace("}", "").Replace("[", "").Replace("]", ""); // to deal with spaces in filenames and invalid values for spark
                string localBlobLocation = Path.Combine(locationToSave, fileName);
                blob.DownloadToFile(localBlobLocation, FileMode.Create);
                return localBlobLocation;
            }
            catch (Exception exception)
            {
                throw new Exception("Exception occured in downloading file from Azure Blob Storage(DownloadBlobFile)\n", exception);
            }
        }

        /// <summary>
        /// check whether given uri is valid by recreating it
        /// </summary>
        /// <param name="inputUri"></param>
        /// <returns>true if valid, false otherwise</returns>
        public bool CheckUriIsValid(string inputUri)
        {
            try
            {
                Uri result;
                if (String.IsNullOrEmpty(inputUri) || String.IsNullOrWhiteSpace(inputUri))
                    return false;

                if (!Uri.TryCreate(inputUri, UriKind.Absolute, out result))
                    return false;

                if (!result.Scheme.Equals(Uri.UriSchemeHttp) && !result.Scheme.Equals(Uri.UriSchemeHttps))
                    return false;

                return true;
            }
            catch (Exception exception)
            {
                throw new Exception("Exception occured in checking blob URI(CheckUriIsValid)\n", exception);
            }
        }

        /// <summary>
        /// save json output to file in local filesystem
        /// </summary>
        /// <param name="jsonData"></param>
        /// <param name="fileName"></param>
        /// <param name="directory"></param>
        /// <returns>saved local file location</returns>
        public string SaveJsonToFile(string jsonData, string fileName, string directory)
        {
      
            if (!iFileSystem.Directory.Exists(directory))
                iFileSystem.Directory.CreateDirectory(directory);

            string outputFileName = Path.Combine(directory, $"{fileName}.json");
            JsonTextWriter jsonTextWriter = new JsonTextWriter(iFileSystem.File.CreateText(outputFileName));
            jsonTextWriter.Close();
            // file is overwritten if already exists
            iFileSystem.File.WriteAllText(outputFileName, jsonData);
            return outputFileName;
        }

        /// <summary>
        /// Convert PDF document to MS Word document
        /// </summary>
        /// <param name="file"></param>
        /// <returns>location of converted/saved word doc</returns>
        public string ConvertPdfToWord(string file, string directoryToSave)
        {
            Document pdfAsWordDoc = null;
            // fire up word instance
            Application wordApp = iInteropWordUtils.CreateWordAppInstance();
            try
            {
                if (!iFileSystem.Directory.Exists(directoryToSave))
                    iFileSystem.Directory.CreateDirectory(directoryToSave);

                pdfAsWordDoc = iInteropWordUtils.OpenDocument(file, wordApp);
                string convertedDocFileLocation = directoryToSave + "/" + Path.ChangeExtension(Path.GetFileName(file), ".doc");
                pdfAsWordDoc.SaveAs2(convertedDocFileLocation, WdSaveFormat.wdFormatDocument);
                return convertedDocFileLocation;
            }
            catch (Exception exception)
            {
                throw new Exception("Exception occured while converting PDF document to Word document (ConvertPdfToWord)\n", exception);
            }
            finally
            {
                // Close without saving and release resources
                pdfAsWordDoc?.Close(SaveChanges: false);
                // Forcefully collect discarded Interop COM objects
                iInteropWordUtils.DisposeIneropObject(wordApp);
            }
        }

        /// <summary>
        /// When Error, prepare json document with error details
        /// </summary>
        /// <param name="fileLocation"></param>
        /// <param name="queueMessage"></param>
        /// <param name="exp"></param>
        /// <returns>Json structure with error details</returns>
        public JsonDocumentStruct PrepareErrorJsonDoc(string fileLocation, Exception exp)
        {
            JsonDocumentStruct jsonDoc = new JsonDocumentStruct();
            jsonDoc.Errors = new Error();
            jsonDoc.Errors.IsError = true;
            jsonDoc.Errors.Description = exp.ToString();
            jsonDoc.FileProperties = ExtractFileMetadata(fileLocation);
            return jsonDoc;
        }

        /// <summary>
        /// Find table containing provided table headers - mainly for document containing tables
        /// </summary>
        /// <param name="wordDocToExtract"></param>
        /// <param name="tableHeaders"></param>
        /// <returns>table index and row index (to know where to start reading data from)</returns>
        public Tuple<int, int> FindTableWithHeader(Document wordDocToExtract, List<string> tableHeaders)
        {
            try
            {
                var tableIndex = 1;
                string cleanColHeader;
                var cellText = new StringBuilder();
                var columnHeadersAsList = new List<string>();
                foreach (Table table in wordDocToExtract.Tables)
                {
                    if (table.Columns.Count == tableHeaders.Count)
                    {
                        // check whether the table contains red flags
                        for (var row = 1; row <= table.Rows.Count; row++)
                        {
                            for (var col = 1; col <= table.Columns.Count; col++)
                            {
                                foreach (Paragraph para in table.Cell(row, col).Range.Paragraphs)
                                    cellText.Append(para.Range.Text);

                                // looks for exact wordings in headers of red flag table
                                cleanColHeader = Regex.Replace(cellText.ToString(), Constants.RegexExp.NoSpecialCharRegex, "");
                                columnHeadersAsList.Add(tableHeaders.Find(x => cleanColHeader.ToLower().Equals(x.ToLower())));
                                cellText.Clear();
                            }

                            // if this table is the table we were looking for, break and return the table index
                            if (!columnHeadersAsList.Contains(null))
                                return Tuple.Create(tableIndex, row + 1);

                            columnHeadersAsList.Clear();
                        }
                    }

                    tableIndex++;
                }

                return Tuple.Create(-1, -1);
            }
            catch (Exception exception)
            {
                throw new Exception("Exception occured on finding table with provided headers in document(FindTableWithHeader)\n", exception);
            }
        }

        /// <summary>
        /// Create cloud blob client from storage account
        /// </summary>
        /// <param name="StorageAccount"></param>
        /// <returns>Cloud blob client</returns>
        public CloudBlobClient CreateCloudBlobClient(CloudStorageAccount StorageAccount)
        {
            return StorageAccount?.CreateCloudBlobClient();
        }

        /// <summary>
        /// get reference for Azure queue from storage account
        /// </summary>
        /// <param name="StorageAccount"></param>
        /// <returns>Queue reference</returns>
        public CloudQueue GetQueueReference(CloudStorageAccount StorageAccount)
        {
            CloudQueueClient queueClient = StorageAccount?.CreateCloudQueueClient();
            return queueClient?.GetQueueReference(CloudConfigurationManager.GetSetting(Constants.ParserConfig.MessageQueueRef));
        }

        /// <summary>
        /// Get Queue message dequeue count
        /// </summary>
        /// <param name="queueMsg"></param>
        /// <returns>message dequeue count</returns>
        public int GetQueueMessageDequeueCount(CloudQueueMessage queueMsg)
        {
            return queueMsg.DequeueCount;
        }
    }
}