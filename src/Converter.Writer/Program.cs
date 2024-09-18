using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Converter.Types;

namespace Converter.Writer
{
    public class Program
    {
        public static byte[] InsertDataToWord(Request request, string filePath)
        {
            byte[] fileBytes = File.ReadAllBytes(filePath);

            using (MemoryStream originalStream = new MemoryStream(fileBytes))
            using (MemoryStream expandableStream = new MemoryStream())
            {
                // Copy the original content to the expandable stream
                originalStream.CopyTo(expandableStream);
                expandableStream.Position = 0; // Reset position to the beginning

                using (WordprocessingDocument doc = WordprocessingDocument.Open(expandableStream, true))
                {
                    InsertChart(request, doc);
                    InsertTable(request, doc);
                    InsertImage(request, doc);
                    InsertText(request, doc);

                    AddUpdateFieldsOnOpen(doc);

                    RemoveAllBookmarks(doc);
                }

                // Return the modified document as a byte array
                return expandableStream.ToArray();
            }
        }

        private static void InsertText(Request request, WordprocessingDocument doc)
        {
            foreach (KeyValuePair<string, string> str in request.Text)
            {
                TextInserter textInserter = new TextInserter(doc);
                textInserter.Insert(str.Key, str.Value);
            }
        }

        private static void InsertImage(Request request, WordprocessingDocument doc)
        {
            if (request.ImageData != null)
            {
                foreach (ImageData imageDataObject in request.ImageData)
                {
                    ImageInserter imageInserter = new ImageInserter(doc);
                    imageInserter.Insert(imageDataObject);
                }
            }
        }

        private static void InsertTable(Request request, WordprocessingDocument doc)
        {
            foreach (TableData tabletDataObject in request.TableData)
            {
                TableInserter tableInserter = new TableInserter(doc);
                tableInserter.Insert(tabletDataObject);
            }
        }

        private static void InsertChart(Request request, WordprocessingDocument doc)
        {
            foreach (ChartData chartDataObject in request.ChartData)
            {
                ChartInserter chartInserter = new ChartInserter(doc);
                chartInserter.Insert(chartDataObject);
            }
        }
        private static void AddUpdateFieldsOnOpen(WordprocessingDocument doc)
        {
            DocumentSettingsPart settingsPart = doc.MainDocumentPart.DocumentSettingsPart;
            if (settingsPart == null)
            {
                settingsPart = doc.MainDocumentPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings();
            }

            UpdateFieldsOnOpen updateFieldsOnOpen = new UpdateFieldsOnOpen
            {
                Val = new OnOffValue(true),
            };

            // Add the new UpdateFields element
            settingsPart.Settings.PrependChild<UpdateFieldsOnOpen>(updateFieldsOnOpen);
            settingsPart.Settings.Save();
        }

        private static void RemoveAllBookmarks(WordprocessingDocument doc)
        {
            RemoveBookmarksFromPart(doc.MainDocumentPart.Document.Body);

            // Remove bookmarks from all headers
            foreach (HeaderPart headerPart in doc.MainDocumentPart.HeaderParts)
            {
                RemoveBookmarksFromPart(headerPart.Header);
            }

            // Remove bookmarks from all footers
            foreach (FooterPart footerPart in doc.MainDocumentPart.FooterParts)
            {
                RemoveBookmarksFromPart(footerPart.Footer);
            }
        }

        private static void RemoveBookmarksFromPart(OpenXmlElement part)
        {
            // Get all BookmarkStart elements in the specified part
            List<BookmarkStart> bookmarks = part.Descendants<BookmarkStart>().ToList();

            foreach (BookmarkStart bookmarkStart in bookmarks)
            {
                string bookmarkId = bookmarkStart.Id.Value;

                // Find corresponding BookmarkEnd
                BookmarkEnd bookmarkEnd = part.Descendants<BookmarkEnd>()
                    .FirstOrDefault(b => b.Id.Value == bookmarkId);

                // Remove BookmarkStart and BookmarkEnd
                if (bookmarkEnd != null)
                {
                    bookmarkStart.Remove();
                    bookmarkEnd.Remove();
                }
            }
        }
    }
}