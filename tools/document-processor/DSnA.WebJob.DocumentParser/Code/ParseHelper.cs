//Copyright(c) Microsoft Corporation.All rights reserved.
//Licensed under the MIT License.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace DSnA.WebJob.DocumentParser
{
    public interface IParseHelper
    {
        DocumentContent ExtractDocumentContent(string docFile, Application wordApp);
    }

    public class ParseHelper : IParseHelper
    {
        internal static string WordHeading1 = "Heading 1";
        internal static string WordHeading2 = "Heading 2";
        internal static string WordHeading3 = "Heading 3";
        internal static string WordHeading4 = "Heading 4";
        private readonly ILogger iLogger;
        private readonly IUtils iUtils;
        private readonly IInteropWordUtils iInteropWordUtils;

        public ParseHelper(ILogger iLogger, IUtils iUtils)
        {
            this.iLogger = iLogger;
            this.iUtils = iUtils;
            iInteropWordUtils = new InteropWordUtils();
        }

        public ParseHelper(ILogger iLogger, IUtils iUtils, IInteropWordUtils iInteropWordUtils)
        {
            this.iLogger = iLogger;
            this.iUtils = iUtils;
            this.iInteropWordUtils = iInteropWordUtils;
        }

        /// <summary>
        /// Extract all paragraphs
        /// </summary>
        /// <param name="wordDocToExtract"></param>
        /// <returns></returns>
        private DocumentContent ExtractAllParagraphs(Document wordDocToExtract, Dictionary<int, string> headers, List<string> tableContent, Dictionary<int, string> listParagraphs)
        {
            try
            {
                var fullContent = string.Empty;
                var paragraphs = new Dictionary<int, string>();
                var sections = new Dictionary<int, string>();

                List<Clauses> clauses = new List<Clauses>();
                List<Clauses> headerClauses = new List<Clauses>();
                List<string> additionalInformation = new List<string>();

                foreach (Paragraph para in wordDocToExtract.Paragraphs)
                {
                    var text = para.Range.Text;
                    var cleanText = iUtils.CleanTextFromNonAsciiChar(text);
                    var textLastSentence = para.Range.Sentences.Last.Text;
                    var textStart = para.Range.Start;
                    fullContent += text;
                    var listNumber = para.Range.ListFormat.ListString;

                    if (!string.IsNullOrEmpty(listNumber))
                    {
                        text = $"{listNumber.Trim()} {text}";
                    }

                    if (textStart > 250 && headers.ContainsKey(textStart))
                    {
                        if (headerClauses.Count > 0)
                        {
                            headerClauses.Last().End = textStart - 1;
                        }

                        headerClauses.Add(new Clauses
                        {
                            Title = headers[textStart],
                            Content = text,
                            Start = textStart
                        });

                        if (headers.Keys.Max() == textStart)
                        {
                            headerClauses.Last().End = para.Range.End;

                            if (!string.IsNullOrEmpty(listNumber))
                            {
                                headerClauses.Last().Content = para.Range.ListFormat.List.Range.Text;
                            }
                        }
                    }
                    else if (headerClauses.Count >= 1 && textStart <= headers.Keys.Max() && tableContent.Contains(cleanText) == false)
                    {
                        headerClauses.Last().Content += text;
                    }

                    if (listParagraphs.Count > 0 && listParagraphs.ContainsKey(textStart))
                    {
                        if(!string.IsNullOrEmpty(cleanText))
                        {
                            sections.Add(textStart, text);
                        }

                        if (clauses.Count > 0)
                        {
                            clauses.Last().End = textStart - 1;
                        }

                        clauses.Add(new Clauses
                        {
                            Title = listParagraphs[textStart],
                            Content = text,
                            Start = textStart
                        });

                        if (listParagraphs.Keys.Max() == textStart)
                        {
                            var nextPara = para.Next();

                            if (nextPara != null)
                            {
                                clauses.Last().End = nextPara.Range.End;
                                clauses.Last().Content += nextPara?.Range?.Text ?? string.Empty;
                            }
                            else
                            {
                                clauses.Last().End = para.Range.End;
                            }
                        }
                    }
                    else if (clauses.Count >= 1 && textStart <= listParagraphs.Keys.Max() && tableContent.Contains(cleanText) == false)
                    {
                        if (!string.IsNullOrEmpty(cleanText))
                        {
                            sections.Add(textStart, text);
                        }

                        clauses.Last().Content += text;
                    }

                    paragraphs.Add(textStart, text);
                }

                if (headerClauses.Count == 0)
                {
                    headerClauses.Add(new Clauses());
                }

                if (clauses.Count == 0)
                {
                    clauses.Add(new Clauses());
                }

                if (sections.Count == 0)
                {
                    sections.Add(-1, string.Empty);
                }

                if (headers.Count == 0)
                {
                    headers.Add(-1, string.Empty);
                }

                StringBuilder rangesContent = new StringBuilder();
                List<Tuple<int, int>> ranges = new List<Tuple<int, int>>();
                List<WdStoryType> rangeTypes = new List<WdStoryType>();

                foreach (Range range in wordDocToExtract.StoryRanges)
                {
                    Range currentRange = range;

                    do
                    {
                        if (RangeStoryTypeIsHeaderOrFooter(currentRange) &&
                            CurrentRangeHaveShapeRanges(currentRange))
                        {
                            foreach (Shape shape in currentRange.ShapeRange)
                            {
                                if (shape.TextFrame.HasText == 0)
                                {
                                    continue;
                                }

                                Range shapeRange = shape.TextFrame.TextRange;

                                rangesContent.Append(RemoveLineBreaks(shapeRange.Text));
                                ranges.Add(new Tuple<int, int>(shapeRange.Start, shapeRange.End));
                                rangeTypes.Add(currentRange.StoryType);
                            }
                        }
                        else
                        {
                            rangesContent.Append(RemoveLineBreaks(currentRange.Text));
                            ranges.Add(new Tuple<int, int>(currentRange.Start, currentRange.End));
                            rangeTypes.Add(currentRange.StoryType);
                        }

                        bool hasMatch = false;
                        MatchCollection matches = Constants.RegexExp.SessionRegEx.Matches(rangesContent.ToString());

                        foreach (Match match in matches)
                        {
                            additionalInformation.Add($"{string.Join("\t", rangeTypes.Select(x => x.ToString()))}|{string.Join("\t", ranges.Select(x => $"{{{x.Item1},{x.Item2}}}"))}|{match.Index}|{match.Value}");
                            hasMatch = true;
                        }

                        matches = Constants.RegexExp.AgendaItemRegEx.Matches(rangesContent.ToString());

                        foreach (Match match in matches)
                        {
                            additionalInformation.Add($"{string.Join("\t", rangeTypes.Select(x => x.ToString()))}|{string.Join("\t", ranges.Select(x => $"{{{x.Item1},{x.Item2}}}"))}|{match.Index}|{match.Value}");
                            hasMatch = true;
                        }

                        if (hasMatch)
                        {
                            rangesContent.Clear();
                            ranges.Clear();
                            rangeTypes.Clear();
                        }

                        currentRange = currentRange.NextStoryRange;
                    } while (currentRange != null);
                }

                return new DocumentContent()
                {
                    Text = fullContent,
                    Paragraphs = paragraphs,
                    Sections = sections,
                    Clauses = clauses,
                    Headers = headers,
                    HeaderClauses = headerClauses,
                    AdditionalInformation = additionalInformation
                };
            }
            catch (Exception exception)
            {
                throw new Exception("Exception in extracting data\n", exception);
            }
        }

        public static string RemoveLineBreaks(string text)
        {
            if (text == "\n"
                || text == "\r\n")
            {
                return " ";
            }

            return text
                .Replace("\r", string.Empty)
                .Replace("\n", string.Empty);
        }

        private static bool RangeStoryTypeIsHeaderOrFooter(Range range)
        {
            return (range.StoryType == WdStoryType.wdEvenPagesHeaderStory ||
                    range.StoryType == WdStoryType.wdPrimaryHeaderStory ||
                    range.StoryType == WdStoryType.wdEvenPagesFooterStory ||
                    range.StoryType == WdStoryType.wdPrimaryFooterStory ||
                    range.StoryType == WdStoryType.wdFirstPageHeaderStory ||
                    range.StoryType == WdStoryType.wdFirstPageFooterStory);
        }

        private static bool CurrentRangeHaveShapeRanges(Range range)
        {
            return range.ShapeRange.Count > 0;
        }

        /// <summary>
        /// Extracts content (red flags, company name, report date) from document
        /// </summary>
        /// <param name="docFile"></param>
        /// <returns>document content</returns>
        public DocumentContent ExtractDocumentContent(string docFile, Application wordApp)
        {
            Document wordDocToExtract = null;

            try
            {
                DocumentContent docContent = new DocumentContent();
                // open the document only in read only mode - so that no edits are made on the document
                wordDocToExtract = iInteropWordUtils.OpenDocument(docFile, wordApp);
                var tableContent = ExtractTableContent(wordDocToExtract);
                // Extract paragraph
                var listParagraphs = ExtractListPragraphs(wordDocToExtract);
                var headers = ExtractHeaders(wordDocToExtract, tableContent, listParagraphs);
                docContent = ExtractAllParagraphs(wordDocToExtract, headers, tableContent, listParagraphs);

                return docContent;
            }
            catch (Exception exception)
            {
                throw new Exception("Exception extracting content (" + nameof(ExtractDocumentContent) + ")\n", exception);
            }
            finally
            {
                // Close without saving and release resources
                wordDocToExtract?.Close(SaveChanges: false);
            }
        }

        private Dictionary<int, string> ExtractListPragraphs(Document wordDocToExtract)
        {
            var listParagraphs = new Dictionary<int, string>();
            foreach (List firstItem in wordDocToExtract.Lists.OfType<List>().Reverse())
            {
                if (firstItem.Range.ListFormat.ListString != null)
                {
                    var totalVlaues = firstItem.Range.ListParagraphs.Count;
                    bool foundNumeric = false;
                    foreach (Paragraph item in firstItem.Range.ListParagraphs.OfType<Paragraph>().Reverse())
                    {
                        if (listParagraphs.ContainsKey(item.Range.Start))
                        {
                            break;
                        }

                        var isNumeric = Regex.IsMatch(item.Range.ListFormat?.ListString ?? string.Empty, Constants.RegexExp.HasNumbers);
                        if (foundNumeric == false)
                        {
                            foundNumeric = isNumeric;
                        }

                        if (foundNumeric == true && isNumeric == false)
                        {
                            continue;
                        }

                        if (item.Range.ListFormat.ListLevelNumber == 1 && (listParagraphs.Count == 0 || listParagraphs.Keys.Max() < item.Range.Start))
                        {
                            listParagraphs.Add(item.Range.Start, item.Range.Sentences.First.Text);
                        }
                    }
                }
            }

            return listParagraphs;
        }

        private List<string> ExtractTableContent(Document wordDocToExtract)
        {
            var tblParaList = new List<string>();
            try
            {
                foreach (Table table in wordDocToExtract.Tables)
                {
                    foreach (Paragraph tblPara in table.Range.Paragraphs)
                    {
                        var cleanText = iUtils.CleanTextFromNonAsciiChar(tblPara.Range.Text).Replace(" ", "");
                        if (!string.IsNullOrEmpty(cleanText))
                        {
                            tblParaList.Add(iUtils.CleanTextFromNonAsciiChar(tblPara.Range.Text));
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            return tblParaList;
        }

        /// <summary>
        /// Extract headers
        /// </summary>
        /// <param name="wordDocToExtract"></param>
        /// <returns>company name</returns>
        private Dictionary<int, string> ExtractHeaders(Document wordDocToExtract, List<string> tblParaList, Dictionary<int, string> listParagraphs)
        {
            try
            {
                var headers = new Dictionary<int, string>();
                foreach (Paragraph para in wordDocToExtract.Paragraphs)
                {
                    try
                    {
                        var textStart = para.Range.Start;
                        if (listParagraphs.Count > 0 && textStart <= listParagraphs.Keys.Max() && textStart >= listParagraphs.Keys.Min())
                        {
                            continue;
                        }

                        var paraText = para.Range.Text;
                        var cleanTextWithoutSpecialChar = iUtils.CleanTextFromNonAsciiChar(Regex.Replace(paraText.ToLower().Trim(), Constants.RegexExp.NoSpecialCharRegex, string.Empty));
                        if (!string.IsNullOrEmpty(cleanTextWithoutSpecialChar) && tblParaList.Contains(iUtils.CleanTextFromNonAsciiChar(paraText)) == false)
                        {
                            string headingStyle = null;
                            try
                            {
                                headingStyle = (para.Range.get_Style() as Style).NameLocal;
                            }
                            catch (Exception)
                            {
                                headingStyle = string.Empty;
                            }

                            if (headingStyle.Equals(WordHeading1) || headingStyle.Equals(WordHeading2) || headingStyle.Equals(WordHeading3) || headingStyle.Equals(WordHeading4) || para.Range.Font.Bold == -1)
                            {
                                if (!Regex.IsMatch(paraText, Constants.RegexExp.OnlyNumericWithSpaces))
                                {
                                    headers.Add(textStart, iUtils.CleanTextFromNonAsciiChar(Regex.Replace(paraText.Replace(".", " ").TrimStart(), Constants.RegexExp.OnlyNumericWithSpaces, string.Empty)));
                                }
                            }
                            else if (para.Range.Words.First.Bold == -1 || para.Range.Font.Size > 12)
                            {
                                var wordCount = para.Range.Sentences.First.Words.Count;
                                if (wordCount <= 6)
                                {
                                    var firstWords = iUtils.CleanTextFromNonAsciiChar(Regex.Replace(para.Range.Sentences.First.Text.Replace(".", " ").TrimStart(), Constants.RegexExp.OnlyNumericWithSpaces, string.Empty));
                                    if (firstWords.Length <= 1)
                                    {
                                        firstWords = iUtils.CleanTextFromNonAsciiChar(Regex.Replace(para.Range.Sentences[2].Text.Replace(".", " ").TrimStart(), Constants.RegexExp.OnlyNumericWithSpaces, string.Empty));
                                    }

                                    headers.Add(textStart, firstWords);
                                }
                                else
                                {
                                    var boldText = string.Empty;
                                    var wordCounter = wordCount <= 25 ? wordCount : 25;
                                    for (int i = 1; i <= wordCount; i++)
                                    {
                                        if (para.Range.Sentences.First.Words[i].Bold == -1)
                                        {
                                            boldText += para.Range.Words[i].Text;
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }

                                    if (boldText != string.Empty)
                                    {
                                        headers.Add(textStart, iUtils.CleanTextFromNonAsciiChar((Regex.Replace(boldText.Replace(".", " ").TrimStart(), Constants.RegexExp.OnlyNumericWithSpaces, string.Empty))));
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                return headers;
            }
            catch (Exception exception)
            {
                throw new Exception("Exception in extracting headers (" + nameof(ExtractHeaders) + ")\n", exception);
            }
        }
    }
}
