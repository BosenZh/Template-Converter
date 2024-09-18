using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using ImageData = Converter.Types.ImageData;
using Justification = DocumentFormat.OpenXml.Wordprocessing.Justification;
using JustificationValues = DocumentFormat.OpenXml.Wordprocessing.JustificationValues;
using NonVisualGraphicFrameDrawingProperties
    = DocumentFormat.OpenXml.Drawing.Wordprocessing.NonVisualGraphicFrameDrawingProperties;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Pic = DocumentFormat.OpenXml.Drawing.Pictures;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;

namespace Converter.Writer
{
    internal class ImageInserter
    {
        private readonly WordprocessingDocument doc;

        public ImageInserter(WordprocessingDocument doc)
        {
            this.doc = doc;
        }

        public void Insert(ImageData imageData)
        {
            this.InsertImageIntoBookmark(imageData.ImageLink, imageData.BookmarkName);
        }

        private static Graphic ObtainGraphic(long cx, long cy, ExternalRelationship imagePart, UInt32Value shapeId)
        {
            Graphic result = new Graphic(
                    new GraphicData(
                        ObtainPicture(cx, cy, imagePart, shapeId))
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" });
            result.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
            return result;
        }

        private static Pic.Picture ObtainPicture(long cx, long cy, ExternalRelationship imagePart, UInt32Value shapeId)
        {
            Pic.Picture result = new Pic.Picture(
                        new Pic.NonVisualPictureProperties(
                            new Pic.NonVisualDrawingProperties() { Id = shapeId, Name = $"Picture {shapeId}" },
                            new Pic.NonVisualPictureDrawingProperties()),
                        new Pic.BlipFill(
                            new Blip() { Link = imagePart.Id, CompressionState = BlipCompressionValues.Print },
                            new Stretch(new FillRectangle())),
                        new Pic.ShapeProperties(
                            new Transform2D(
                                new Offset() { X = 0L, Y = 0L },
                                new Extents() { Cx = cx, Cy = cy }),
                            new PresetGeometry(new AdjustValueList()) { Preset = ShapeTypeValues.Rectangle }));
            result.AddNamespaceDeclaration("pic", "http://schemas.openxmlformats.org/drawingml/2006/picture");
            return result;
        }

        private T GetOrCreatePrependedChild<T>(OpenXmlCompositeElement parent)
            where T : OpenXmlElement, new()
        {
            T child = parent.GetFirstChild<T>();
            if (child == null)
            {
                child = new T();
                parent.PrependChild(child);
            }

            return child;
        }

        private void InsertImageIntoBookmark(string imageLink, string bookmarkName)
        {
            MainDocumentPart mainPart = this.doc.MainDocumentPart;
            Body body = mainPart.Document.Body;
            ExternalRelationship imagePart = mainPart.AddExternalRelationship(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                new Uri(imageLink, UriKind.Absolute));

            Random random = new Random();
            UInt32Value shapeId = new UInt32Value((uint)random.Next(1, int.MaxValue));

            BookmarkStart bookmarkStart = body.Descendants<BookmarkStart>()
                                              .FirstOrDefault(b => b.Name == bookmarkName);
            BookmarkEnd bookmarkEnd = body.Descendants<BookmarkEnd>()
                                          .FirstOrDefault(b => b.Id == bookmarkStart.Id);

            Paragraph parentParagraph = bookmarkStart.Ancestors<Paragraph>().FirstOrDefault();

            // Obtain the original run that contains the bookmark
            this.SetNewRunProperties(parentParagraph, out long cx, out long cy, out Justification justification);

            GraphicFrameLocks graphicFrameLocks = new GraphicFrameLocks() { NoChangeAspect = true };
            graphicFrameLocks.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            Drawing newDrawing = this.GetDrawing(imagePart, shapeId, cx, cy, graphicFrameLocks);
            Run newRun = new Run(new RunProperties(new NoProof()), newDrawing);
            if (parentParagraph != null)
            {
                parentParagraph.InsertBefore(newRun, bookmarkEnd);
            }
            else
            {
                parentParagraph = new Paragraph(newRun);

                body.InsertBefore(parentParagraph, bookmarkEnd);
            }

            ParagraphProperties pPr = this.GetOrCreatePrependedChild<ParagraphProperties>(parentParagraph);
            ParagraphStyleId paragraphStyleId = new ParagraphStyleId() { Val = "Normal" };
            pPr.Append(justification, paragraphStyleId);
        }

        private void SetNewRunProperties(Paragraph parentParagraph, out long cx, out long cy, out Justification justification)
        {
            Run originalRun = parentParagraph?.Descendants<Run>().FirstOrDefault(r => r.Descendants<Drawing>().Any());
            cx = 0;
            cy = 0;
            justification = new Justification() { Val = JustificationValues.Left };
            if (originalRun != null)
            {
                Drawing originalDrawing = originalRun.Descendants<Drawing>().FirstOrDefault();
                Extent extent = originalDrawing.Inline.Extent;
                cx = extent.Cx;
                cy = extent.Cy;

                originalRun.Remove();
            }
            else
            {
                cx = 3000000; // Default width (in EMUs)
                cy = 2000000; // Default height (in EMUs)
                justification = new Justification() { Val = JustificationValues.Center };
            }
        }

        private Drawing GetDrawing(ExternalRelationship imagePart, UInt32Value shapeId, long cx, long cy, GraphicFrameLocks graphicFrameLocks)
        {
            return new Drawing(
                            new Inline(
                                new Extent() { Cx = cx, Cy = cy },
                                new EffectExtent() { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                                new DocProperties() { Id = shapeId, Name = $"Picture {shapeId}" },
                                new NonVisualGraphicFrameDrawingProperties(graphicFrameLocks),
                                ObtainGraphic(cx, cy, imagePart, shapeId)));
        }
    }
}
