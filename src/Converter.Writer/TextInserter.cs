using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Converter.Writer
{
    internal class TextInserter
    {
        private readonly WordprocessingDocument doc;

        public TextInserter(WordprocessingDocument doc)
        {
            this.doc = doc;
        }

        public void Insert(string bookmarkName, string replaceText)
        {
            this.WriteTextToBookmark(bookmarkName, replaceText);
        }

        private void ReplaceBookmarkText(OpenXmlPart part, string bookmarkName, string replaceText)
        {
            const string whiteHex = "FFFFFF";
            IEnumerable<BookmarkStart> bookmarks = part.RootElement.Descendants<BookmarkStart>();

            foreach (BookmarkStart bookmarkStart in bookmarks)
            {
                if (bookmarkStart.Name != bookmarkName)
                {
                    continue;
                }

                OpenXmlElement currentElement = bookmarkStart.NextSibling();
                BookmarkEnd bookmarkEnd = part.RootElement.Descendants<BookmarkEnd>()
                    .FirstOrDefault(b => b.Id == bookmarkStart.Id);

                if (bookmarkEnd == null)
                {
                    Debug.WriteLine($"{bookmarkName} does not have a corresponding BookmarkEnd");
                    continue;
                }

                RunProperties copiedRunProperties = null;
                OpenXmlElement siblingElement = this.GetSiblingElement(bookmarkStart, ref copiedRunProperties);

                // If no previous sibling Run found, check next siblings
                if (copiedRunProperties == null)
                {
                    siblingElement = this.CheckNextSibling(bookmarkStart, ref copiedRunProperties);
                }

                Paragraph parentParagraph = this.GetparentParagraph(bookmarkStart);

                Run newRun = new Run();

                if (!string.IsNullOrEmpty(replaceText) && replaceText.StartsWith("\r"))
                {
                    newRun.AppendChild(new Break());
                    newRun.AppendChild(new Break());
                }

                newRun.AppendChild(new Text(replaceText));

                // Apply the copied RunProperties if available
                this.SetCopiedRunProperties(bookmarkName, whiteHex, copiedRunProperties, newRun);

                parentParagraph.InsertAfter(newRun, bookmarkStart);

                // Remove the elements between the BookmarkStart and BookmarkEnd
                while (currentElement != null && currentElement != bookmarkEnd)
                {
                    OpenXmlElement temp = currentElement.NextSibling();
                    currentElement.Remove();
                    currentElement = temp;
                }
            }
        }

        private void SetCopiedRunProperties(string bookmarkName, string whiteHex, RunProperties copiedRunProperties, Run newRun)
        {
            if (copiedRunProperties != null)
            {
                newRun.PrependChild(copiedRunProperties.CloneNode(true) as RunProperties);
            }
            else if (bookmarkName == "D1")
            {
                newRun.PrependChild(
                    new RunProperties(
                        new Color() { Val = whiteHex }));
            }
        }

        private OpenXmlElement CheckNextSibling(BookmarkStart bookmarkStart, ref RunProperties copiedRunProperties)
        {
            OpenXmlElement siblingElement = bookmarkStart.NextSibling();
            while (siblingElement != null && copiedRunProperties == null)
            {
                if (siblingElement is Run run)
                {
                    copiedRunProperties = run.GetFirstChild<RunProperties>()?.CloneNode(true) as RunProperties;
                }

                siblingElement = siblingElement.NextSibling();
            }

            return siblingElement;
        }

        private Paragraph GetparentParagraph(BookmarkStart bookmarkStart)
        {
            if (!(bookmarkStart.Parent is Paragraph parentParagraph))
            {
                parentParagraph = bookmarkStart.Ancestors<Paragraph>().FirstOrDefault();
            }

            if (parentParagraph == null)
            {
                parentParagraph = new Paragraph();
                OpenXmlElement parentElement = bookmarkStart.Ancestors().FirstOrDefault(
                    a => a is Body
                    || a is TableCell
                    || a is SdtBlock);
                if (parentElement != null)
                {
                    parentElement.InsertBefore(parentParagraph, bookmarkStart);
                    bookmarkStart.Parent?.RemoveChild(bookmarkStart);
                    parentParagraph.AppendChild(bookmarkStart);
                }
            }

            return parentParagraph;
        }

        private OpenXmlElement GetSiblingElement(BookmarkStart bookmarkStart, ref RunProperties copiedRunProperties)
        {
            OpenXmlElement siblingElement = bookmarkStart.PreviousSibling();
            while (siblingElement != null && copiedRunProperties == null)
            {
                if (siblingElement is Run run)
                {
                    copiedRunProperties = run.GetFirstChild<RunProperties>()?.CloneNode(true) as RunProperties;
                }

                siblingElement = siblingElement.PreviousSibling();
            }

            return siblingElement;
        }

        private void WriteTextToBookmark(string bookmarkName, string replaceText)
        {
            MainDocumentPart mainPart = this.doc.MainDocumentPart;
            if (mainPart != null)
            {
                this.ReplaceBookmarkText(mainPart, bookmarkName, replaceText);
            }

            foreach (HeaderPart headerPart in this.doc.MainDocumentPart.HeaderParts)
            {
                this.ReplaceBookmarkText(headerPart, bookmarkName, replaceText);
            }

            foreach (FooterPart footerPart in this.doc.MainDocumentPart.FooterParts)
            {
                this.ReplaceBookmarkText(footerPart, bookmarkName, replaceText);
            }
        }
    }
}
