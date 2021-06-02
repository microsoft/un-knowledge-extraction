//Copyright(c) Microsoft Corporation.All rights reserved.
//Licensed under the MIT License.

using System;
using System.IO;
using Microsoft.WindowsAzure.Storage.Blob;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace DSnA.WebJob.DocumentParser
{
    public interface IDocumentParser
    {
        string ParseDocuments(string blobUri, string outputFileFormat, CloudBlobClient blobClient);
    }
    public class DocumentParser : IDocumentParser
    {
        private readonly ILogger iLogger;
        private readonly IUtils iUtils;
        private readonly IParseHelper iparseHelper;

        public DocumentParser(ILogger iLogger, IUtils iUtils, IParseHelper iparseHelper)
        {
            this.iLogger = iLogger;
            this.iUtils = iUtils;
            this.iparseHelper = iparseHelper;
        }

        /// <summary>
        /// Main API for document extraction
        /// </summary>
        public string ParseDocuments(string blobUri, string outputFileFormat, CloudBlobClient blobClient)
        {
            try
            {
                string fileLocation = "";
                var output = "";
                try
                {
                    fileLocation = iUtils.DownloadBlobFile(blobUri, Constants.FileConfigs.WorkingDirectoryPath, blobClient);
                    var result = ExtractContentFromReports(fileLocation, outputFileFormat, blobUri);
                    iUtils.UploadFileToBlob(result.location, blobClient);
                    output = $"Finished Processing: {result.location}";
                    iUtils.DeleteInputFiles(new List<string> { result.location });
                }
                catch (Exception exp)
                {
                    iLogger.Error($"Error Processing: {blobUri}", exp);
                    output = $"error processing: {blobUri}";
                }

                return output;
            }
            catch (Exception exp)
            {
                iLogger.Error("{" + nameof(ParseDocuments) + "} - exception occured-Level 2", exp);
                throw;
            }
            finally
            {
                // force garbage collection to collect leftover COM objects
                GC.Collect();
            }
        }

        /// <summary>
        /// Extracts document content - Initial function encapsulating different extraction helper methods
        /// </summary>
        /// <param name="fileLocation"></param>
        /// <param name="queueMessage"></param>
        /// <returns>saved JSON output file location</returns>
        private ReportExtractionResponse ExtractContentFromReports(string fileLocation, string outputFileFormat, string blobUri = null)
        {
            var docFile = "";
            try
            {
                docFile = iUtils.ConvertPdfToWord(fileLocation, Constants.FileConfigs.WorkingDirectoryPath);
                var documentContent = iparseHelper.ExtractDocumentContent(docFile);
                ReportExtractionResponse reportExtractionResponse = null;
                switch (outputFileFormat)
                {
                    case Constants.CsvFileConfig.JsonFileFormat:
                        JsonDocumentStructFlat jsonDoc;
                        string jsonOutputFileLocation;
                        ExtractAsJsonFormat(fileLocation, blobUri, documentContent, out jsonDoc, out jsonOutputFileLocation);
                        reportExtractionResponse = new ReportExtractionResponse()
                        {
                            location = jsonOutputFileLocation,
                            contentJson = JsonConvert.SerializeObject(jsonDoc, Formatting.Indented)
                        };
                        break;

                    case Constants.CsvFileConfig.CsvFileFormat:
                    default:
                        string csvOutputFileLocation;
                        ExtractAsCsvFormat(fileLocation, blobUri, documentContent, out csvOutputFileLocation);
                        reportExtractionResponse = new ReportExtractionResponse()
                        {
                            location = csvOutputFileLocation
                        };
                        break;
                }

                return reportExtractionResponse;
            }
            finally
            {
                // delete input files - leave no trace of confidential documents
                iUtils.DeleteInputFiles(new List<string> { fileLocation, docFile });
            }
        }

        private void ExtractAsJsonFormat(string fileLocation, string blobUri, DocumentContent documentContent, out JsonDocumentStructFlat jsonDoc, out string jsonOutputFileLocation)
        {
            jsonDoc = new JsonDocumentStructFlat();
            jsonDoc.ImageStoreUri = blobUri;
            jsonDoc.Text = documentContent.Text;
            jsonDoc.Paragraphs = documentContent.Paragraphs;
            jsonDoc.Headers = documentContent.Headers;
            jsonDoc.Sections = documentContent.Sections;
            jsonDoc.Clauses = documentContent.Clauses;
            jsonDoc.HeaderClauses = documentContent.HeaderClauses;
            var fileProperties = iUtils.ExtractFileMetadata(fileLocation);
            jsonDoc.FileName = fileProperties.FileName;
            jsonDoc.FileType = fileProperties.FileType;
            jsonDoc.AgreementNumber = fileProperties.AgreementNumber;
            jsonDoc.ExtractionTimeStamp = fileProperties.ExtractionTimeStamp;
            jsonOutputFileLocation = iUtils.SerializeAndSaveJson(jsonDoc, Path.GetFileName(fileLocation));
        }

        private void ExtractAsCsvFormat(string fileLocation, string blobUri, DocumentContent documentContent, out string csvOutputFileLocation)
        {
            CsvDocumentFile csvDocumentFile = new CsvDocumentFile();
            var fileProperties = iUtils.ExtractFileMetadata(fileLocation);
            csvDocumentFile.AddCsvLine(fileProperties.FileName, blobUri, Constants.CsvFileConfig.ContentTypeBlobUri);
            csvDocumentFile.AddCsvLine(fileProperties.FileName, fileProperties.AgreementNumber, Constants.CsvFileConfig.ContentTypeAgreementNumber);
            csvDocumentFile.AddCsvLine(fileProperties.FileName, fileProperties.FileType, Constants.CsvFileConfig.ContentTypeFileType);
            csvDocumentFile.AddCsvLine(fileProperties.FileName, fileProperties.ExtractionTimeStamp, Constants.CsvFileConfig.ContentTypeExtractionTimeStamp);
            csvDocumentFile.AddCsvLine(fileProperties.FileName, iUtils.CleanTextFromNonAsciiChar(documentContent.Text), Constants.CsvFileConfig.ContentTypeText);

            foreach (var paragraph in documentContent.Paragraphs)
            {
                var paragraphCleanText = iUtils.CleanTextFromNonAsciiChar(paragraph.Value);
                if (!string.IsNullOrEmpty(paragraphCleanText))
                    csvDocumentFile.AddCsvLine(fileProperties.FileName, paragraphCleanText, Constants.CsvFileConfig.ContentTypeParagraph);
            }

            foreach (var header in documentContent.Headers)
            {
                var headerCleanText = iUtils.CleanTextFromNonAsciiChar(header.Value);
                if (!string.IsNullOrEmpty(headerCleanText))
                    csvDocumentFile.AddCsvLine(fileProperties.FileName, headerCleanText, Constants.CsvFileConfig.ContentTypeHeader);
            }

            foreach (var section in documentContent.Sections)
            {
                var sectionCleanText = iUtils.CleanTextFromNonAsciiChar(section.Value);
                if (!string.IsNullOrEmpty(sectionCleanText))
                    csvDocumentFile.AddCsvLine(fileProperties.FileName, sectionCleanText, Constants.CsvFileConfig.ContentTypeSection);
            }

            foreach (var clause in documentContent.Clauses)
            {
                var clauseCleanText = iUtils.CleanTextFromNonAsciiChar(clause.Content);
                if (!string.IsNullOrEmpty(clauseCleanText))
                    csvDocumentFile.AddCsvLine(fileProperties.FileName, clauseCleanText, Constants.CsvFileConfig.ContentTypeClause);
            }

            foreach (var headerClause in documentContent.HeaderClauses)
            {
                var headerClauseCleanText = iUtils.CleanTextFromNonAsciiChar(headerClause.Content);
                if (!string.IsNullOrEmpty(headerClauseCleanText))
                    csvDocumentFile.AddCsvLine(fileProperties.FileName, headerClauseCleanText, Constants.CsvFileConfig.ContentTypeHeaderClause);
            }

            csvOutputFileLocation = iUtils.SaveToCsvFile(csvDocumentFile.GetCsvOutputLines(), Path.GetFileNameWithoutExtension(fileLocation));
        }
    }
}