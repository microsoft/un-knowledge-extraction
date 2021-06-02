//Copyright(c) Microsoft Corporation.All rights reserved.
//Licensed under the MIT License.

namespace DSnA.WebJob.DocumentParser
{
    using System;
    public static class Constants
    {
        public static string DateTimeFormat => "MM-dd-yyyy_HH-mm-ss";
        public static class FileConfigs
        {
            public static string SourceDirectoryPath => @"C:\data\DocParsing";// @"%HOME%\data\";//
            public static string WorkingDirectoryPath => @"C:\data\DocParsing\Temp";//  @"%HOME%\data\tempForProcessing";//
            public static string OutputDirectoryPath => SourceDirectoryPath + @"\Output";
            public static string LogFileName = "log_" + DateTime.UtcNow.ToString("MM-dd-yyyy") + ".log";
            public static string TempFileName => "JsonByExtractionProgram";
        }

        public static class RegexExp
        {
            public static string NoSpecialCharRegex => "[\\W]+";
            // match only strings with combination of numbers and spaces
            public static string OnlyNumericWithSpaces => "^([0-9\\s]+)$";
            // to match company names in reports
            public static string CompanyNameRegex => "^(^[a-zA-Z\\d\\s]+[a-zA-Z\\d]+[a-zA-Z\\d\\W]*)$";
            // regex to match dates like MMMM dd,YYYY (January 23, 2017) and its combinations
            public static string DateRegex => "^(\\s*\\w{3,9}?\\s*?\\d{1,2}?\\s*?,\\s*?\\d{4}?)";
            public static string NoEscapeSequences => @"[\a\b\t\r\v\f\e]";
            public static string OnlyAsciiChar => @"[^\u0000-\u007F]+";
            public static string OnlyWhiteSpaces => @"\s+";
            public static string OnlyHyperlinks => @"(http|ftp|https)://([\w_-]+(?:(?:\.[\w_-]+)+))([\w.,@?^=%&:/~+#-]*[\w@?^=%&/~+#-])?";
            public static string HyperlinkAppProtocols => @"(http|https|ftp)";
            public static string HasNumbers => @"^(?=.*[0-9])";

            public static string HasBulletPoint => @"^(|\u2022|\u2023|\u25E6|\u2043|\u2219|-|[a-z]\)|[a-z]\.)";
        }

        public static class ParserConfig
        {
            public static string MessageQueueRef => "QueueName";
            public static string ConnectionUriRef => "KeyVaultUriForConnectionString";
            public static int MaxDequeueCount => 5;
            public static string LogsContainerNameRef => "LogsContainerName";
            public static string LogPrefix => "LogPrefix";
            public static string OutputContainerNameRef => "OutputContainerName";
        }

        public static class CsvFileConfig
        {
            public const string CsvFileFormat = "csv";
            public const string JsonFileFormat = "json";
            public static string Headers => "SourceFile,Index,Content,Type";
            public static string ContentTypeBlobUri => "BlobUri";
            public static string ContentTypeAgreementNumber => "AgreementNumber";
            public static string ContentTypeFileType => "FileType";
            public static string ContentTypeExtractionTimeStamp => "ExtractionTimeStamp";
            public static string ContentTypeText => "Text";
            public static string ContentTypeParagraph => "Paragraph";
            public static string ContentTypeHeader => "Header";
            public static string ContentTypeSection => "Section";
            public static string ContentTypeClause => "Clause";
            public static string ContentTypeHeaderClause => "HeaderClause";
        }
    }
}