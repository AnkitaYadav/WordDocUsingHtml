using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using PreMailer.Net;
using HandlebarsDotNet;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using NotesFor.HtmlToOpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Linq;

namespace HtmlToDocUsingHandlerBar
{
    class Program
    {
        static void Main(string[] args)
        {
            var client = new WebClient();
            var meetingResponse = client.DownloadString("http://38.118.71.177/nyacuradealers/service//meeting/meetingDetails?meetingId=12&associationId=1");
            var attendeesResponse = client.DownloadString("http://38.118.71.177/nyacuradealers/service/meeting/attendees?meetingId=12");
            var htmlTemplate = File.ReadAllText(@"E:\Ankita\Project\ConversionFromHTMLtoDOC\HtmlToDocUsingHandlerBar\WordDocument.html");
            JavaScriptSerializer serialiser = new JavaScriptSerializer();
            var meetingInfo = serialiser.Deserialize<MeetingModel>(meetingResponse);
            var attendees = serialiser.Deserialize<IEnumerable<Attendee>>(attendeesResponse);
            meetingInfo.Attendees = attendees;
            var html = PreMailer.Net.PreMailer.MoveCssInline(htmlTemplate).Html;
            var tempalte = Handlebars.Compile(html);
            var result = tempalte(meetingInfo);
            using (MemoryStream generatedDocument = new MemoryStream())
            {
                using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
                {
                    MainDocumentPart mainPart = package.MainDocumentPart;
                    if (mainPart == null)
                    {
                        mainPart = package.AddMainDocumentPart();
                        var pageHeaderPart = mainPart.AddNewPart<HeaderPart>("rId2");
                        string fileName = @"E:\Ankita\Project\ConversionFromHTMLtoDOC\DocWithHeader\GCHDA.png";
                        
                        GeneratePageHeaderPart("Quaterly Meeting ").Save(pageHeaderPart);
                        foreach (var header in package.MainDocumentPart.HeaderParts)
                        {
                            Header hd = header.Header;

                            ImagePart imagePart = header.AddImagePart(ImagePartType.Jpeg);

                            using (FileStream stream = new FileStream(fileName, FileMode.Open))
                            {
                                imagePart.FeedData(stream);
                            }
                            AddImageToBody(hd, header.GetIdOfPart(imagePart));
                        }

                        new Document(new Body(
                             new SectionProperties(
                        new HeaderReference()
                        {
                            Type = HeaderFooterValues.Default,
                            Id = "rId2"
                        }
                               )
                            )).Save(mainPart);
                    }

                    HtmlConverter converter = new HtmlConverter(mainPart);
                    converter.ParseHtml(result);

                    mainPart.Document.Save();
                }
                File.WriteAllBytes("E://testWord12.docx", generatedDocument.ToArray());
            }
        }
        private static Header GeneratePageHeaderPart(string HeaderText)
        {
            var element =
                new Header(
                    new Paragraph(
                        new ParagraphProperties(
                            new ParagraphStyleId() { Val = "Header" },
                           new Justification() { Val = JustificationValues.Right }
                            ),
                        new Run(
                            new Text(HeaderText))
                    ));


            return element;
        }
        private static void AddImageToBody(Header hd, string relationshipId)
        {
            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = 990000L, Cy = 792000L },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                       "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            // Append the reference to the header, the element should be in a Run.
            hd.AppendChild(new Paragraph(
                            new ParagraphProperties(
                           new Justification() { Val = JustificationValues.Left }
                            ),
                               new Run(element)
                               ));
        }
    }
}
