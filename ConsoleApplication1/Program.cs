using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        string documentPath = @"E:\\testWord12.docx";
        string fileName = @"E:\Ankita\Project\ConversionFromHTMLtoDOC\DocWithHeader\GCHDA.png";
        using (WordprocessingDocument wordprocessingDocument =
                   WordprocessingDocument.Open(documentPath, true))
        {
            MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;
            HeaderPart hdPart = mainPart.GetPartsOfType<HeaderPart>().FirstOrDefault();
            Header hd = hdPart.Header;

            ImagePart imagePart = hdPart.AddImagePart(ImagePartType.Jpeg);

            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }
            AddImageToBody(hd, hdPart.GetIdOfPart(imagePart));
        }

    }
    //public static void InsertAPicture(string document, string fileName)
    //{
    //    using (WordprocessingDocument wordprocessingDocument =
    //        WordprocessingDocument.Open(document, true))
    //    {
           
           
    //        MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;
    //        foreach (var header in wordprocessingDocument.MainDocumentPart.HeaderParts)
    //        {
    //            ImagePart imagePart = header.AddImagePart(ImagePartType.Jpeg);
    //            using (FileStream stream = new FileStream(fileName, FileMode.Open))
    //            {
    //                imagePart.FeedData(stream);
    //            }
    //            AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
    //        }

    //        //ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

          

           
    //    }
    //}
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
        hd.AppendChild(new Paragraph(new Run(element)));
    }

    //private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
    //{
    //    // Define the reference of the image.
    //    var element =
    //         new Drawing(
    //             new DW.Inline(
    //                 new DW.Extent() { Cx = 990000L, Cy = 792000L },
    //                 new DW.EffectExtent()
    //                 {
    //                     LeftEdge = 0L,
    //                     TopEdge = 0L,
    //                     RightEdge = 0L,
    //                     BottomEdge = 0L
    //                 },
    //                 new DW.DocProperties()
    //                 {
    //                     Id = (UInt32Value)1U,
    //                     Name = "Picture 1"
    //                 },
    //                 new DW.NonVisualGraphicFrameDrawingProperties(
    //                     new A.GraphicFrameLocks() { NoChangeAspect = true }),
    //                 new A.Graphic(
    //                     new A.GraphicData(
    //                         new PIC.Picture(
    //                             new PIC.NonVisualPictureProperties(
    //                                 new PIC.NonVisualDrawingProperties()
    //                                 {
    //                                     Id = (UInt32Value)0U,
    //                                     Name = "New Bitmap Image.jpg"
    //                                 },
    //                                 new PIC.NonVisualPictureDrawingProperties()),
    //                             new PIC.BlipFill(
    //                                 new A.Blip(
    //                                     new A.BlipExtensionList(
    //                                         new A.BlipExtension()
    //                                         {
    //                                             Uri =
    //                                                "{28A0092B-C50C-407E-A947-70E740481C1C}"
    //                                         })
    //                                 )
    //                                 {
    //                                     Embed = relationshipId,
    //                                     CompressionState =
    //                                     A.BlipCompressionValues.Print
    //                                 },
    //                                 new A.Stretch(
    //                                     new A.FillRectangle())),
    //                             new PIC.ShapeProperties(
    //                                 new A.Transform2D(
    //                                     new A.Offset() { X = 0L, Y = 0L },
    //                                     new A.Extents() { Cx = 990000L, Cy = 792000L }),
    //                                 new A.PresetGeometry(
    //                                     new A.AdjustValueList()
    //                                 )
    //                                 { Preset = A.ShapeTypeValues.Rectangle }))
    //                     )
    //                     { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
    //             )
    //             {
    //                 DistanceFromTop = (UInt32Value)0U,
    //                 DistanceFromBottom = (UInt32Value)0U,
    //                 DistanceFromLeft = (UInt32Value)0U,
    //                 DistanceFromRight = (UInt32Value)0U,
    //                 EditId = "50D07946"
    //             });

    //    // Append the reference to body, the element should be in a Run.
    //    //wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
    //}
    //private static Header GeneratePageHeaderPart(string headerText)
    //{
    //    //Header hdr = new Header(new Paragraph(new Run(LoadImage(_agLogoRel, _agLogoFilename, "name" + _agLogoRel, 2.57, 0.73))));
    //    //return hdr;
    //}
    //private static Drawing LoadImage(string relationshipId,
    //                         string filename,
    //                         string picturename,
    //                         double inWidth,
    //                         double inHeight)
    //{
    //    //double emuWidth = Konsts.EmusPerInch * inWidth;
    //    //double emuHeight = Konsts.EmusPerInch * inHeight;

    //    var element = new Drawing(
    //        new DW.Inline(
    //       // new DW.Extent { Cx = (Int64Value)emuWidth, Cy = (Int64Value)emuHeight },
    //        new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
    //        new DW.DocProperties { Id = (UInt32Value)1U, Name = picturename },
    //        new DW.NonVisualGraphicFrameDrawingProperties(
    //        new A.GraphicFrameLocks { NoChangeAspect = true }),
    //        new A.Graphic(
    //        new A.GraphicData(
    //        new PIC.Picture(
    //        new PIC.NonVisualPictureProperties(
    //        new PIC.NonVisualDrawingProperties { Id = (UInt32Value)0U, Name = filename },
    //        new PIC.NonVisualPictureDrawingProperties()),
    //        new PIC.BlipFill(
    //        new A.Blip(
    //        new A.BlipExtensionList(
    //        new A.BlipExtension { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" }))
    //        {
    //            Embed = relationshipId,
    //            CompressionState = A.BlipCompressionValues.Print
    //        },
    //        new A.Stretch(
    //        new A.FillRectangle())),
    //        new PIC.ShapeProperties(
    //        new A.Transform2D(
    //        new A.Offset { X = 0L, Y = 0L },
    //       // new A.Extents { Cx = (Int64Value)emuWidth, Cy = (Int64Value)emuHeight }),
    //        new A.PresetGeometry(
    //        new A.AdjustValueList())
    //        { Preset = A.ShapeTypeValues.Rectangle })))
    //        {
    //            Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
    //        }))
    //       );
    //    return element;
    //}

}