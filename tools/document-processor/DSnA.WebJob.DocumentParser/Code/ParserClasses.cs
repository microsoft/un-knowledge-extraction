//Copyright(c) Microsoft Corporation.All rights reserved.
//Licensed under the MIT License.

namespace DSnA.WebJob.DocumentParser
{
    using Newtonsoft.Json;
    using System.Collections.Generic;

    /// <summary>
    /// Represents output JSON document structure
    /// </summary>
    public class JsonDocumentStruct
    {
        public DocumentContent DocumentContent { get; set; }
        public FileMetaData FileProperties { get; set; }
        public Error Errors { get; set; }
    }

    public class JsonDocumentStructFlat
    {
        [JsonProperty(PropertyName = "agreementNumber")]
        public string AgreementNumber { get; set; }

        [JsonProperty(PropertyName = "fileName")]
        public string FileName { get; set; }

        [JsonProperty(PropertyName = "fileType")]
        public string FileType { get; set; }

        [JsonProperty(PropertyName = "imageStoreUri")]
        public string ImageStoreUri { get; set; }

        [JsonProperty(PropertyName = "extractionTimeStamp")]
        public string ExtractionTimeStamp { get; set; }

        [JsonProperty(PropertyName = "text")]
        public string Text { get; set; }

        [JsonProperty(PropertyName = "headers")]
        public Dictionary<int, string> Headers { get; set; }

        [JsonProperty(PropertyName = "paragraphs")]
        public Dictionary<int, string> Paragraphs { get; set; }

        [JsonProperty(PropertyName = "sections")]
        public Dictionary<int, string> Sections { get; set; }

        [JsonProperty(PropertyName = "clauses")]
        public List<Clauses> Clauses { get; set; }

        [JsonProperty(PropertyName = "headerClauses")]
        public List<Clauses> HeaderClauses { get; set; }
    }

    public class ReportExtractionResponse
    {
        public string location { get; set; }
        public string contentJson { get; set; }
    }

    public class Clauses
    {
        public Clauses()
        {
            this.Title = "";
            this.Content = "";
            this.Start = -1;
            this.End = -1;
        }

        public string Title { get; set; }
        public string Content { get; set; }
        public int Start { get; set; }
        public int End { get; set; }
    }

    /// <summary>
    /// Represents file meta data in Json output
    /// </summary>
    public class FileMetaData
    {
        [JsonProperty(PropertyName = "agreementNumber")]
        public string AgreementNumber { get; set; }

        [JsonProperty(PropertyName = "fileName")]
        public string FileName { get; set; }

        [JsonProperty(PropertyName = "fileType")]
        public string FileType { get; set; }

        [JsonProperty(PropertyName = "extractionTimeStamp")]
        public string ExtractionTimeStamp { get; set; }
    }

    /// <summary>
    /// Represents higher structure of red flag document content in Json output
    /// </summary>
    public class DocumentContent
    {
        [JsonProperty(PropertyName = "text")]
        public string Text { get; set; }
        [JsonProperty(PropertyName = "paragraphs")]
        public Dictionary<int, string> Paragraphs { get; set; }

        [JsonProperty(PropertyName = "headers")]
        public Dictionary<int, string> Headers { get; set; }

        [JsonProperty(PropertyName = "sections")]
        public Dictionary<int, string> Sections { get; set; }

        [JsonProperty(PropertyName = "clauses")]
        public List<Clauses> Clauses { get; set; }

        [JsonProperty(PropertyName = "headerClauses")]
        public List<Clauses> HeaderClauses { get; set; }
    }

    /// <summary>
    /// Represents structure of errors presented in Json output
    /// </summary>
    public class Error
    {
        public Error()
        {
            this.IsError = false;
            this.Description = "";
        }

        public bool IsError { get; set; }
        public string Description { get; set; }
    }

    /// <summary>
    /// Represents structure of input message from OneVet Queue
    /// </summary>
    public class QueueMessage
    {
        public string DocumentId { get; set; }
        public string FileInputUri { get; set; }
        public string FileOutputUri { get; set; }
        public string RequestCreationDateTimeUtc { get; set; }
        public string DocumentTypeId { get; set; }
    }

    /// <summary>
    /// Represents structure for csv output files
    /// </summary>
    public class CsvDocumentFile
    {
        public CsvDocumentFile()
        {
            this.CsvOutputLines = new List<string>() { Constants.CsvFileConfig.Headers };
        }

        private List<string> CsvOutputLines { get; set; }

        public void AddCsvLine(string sourceFile, string content, string type)
        {
            this.CsvOutputLines.Add(string.Format("{0},{1},\"{2}\",{3}", sourceFile, GetCurrentLineCount(), content, type));
        }

        public List<string> GetCsvOutputLines()
        {
            return this.CsvOutputLines;
        }

        private int GetCurrentLineCount()
        {
            return this.CsvOutputLines.Count > 0 ? this.CsvOutputLines.Count - 1 : this.CsvOutputLines.Count;
        }
    }
}
