using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;

namespace ImageHeader
{
    class Program
    {
        static void Main(string[] args)
        {
            //Load Document
            Document doc = new Document();
            doc.LoadFromFile(@"E://testWord12.docx");

            //Header Paragraph
            HeaderFooter header = doc.Sections[0].HeadersFooters.Header;
            Paragraph para = header.AddParagraph();
            para.Format.HorizontalAlignment = HorizontalAlignment.Left;

            //Header Image
            DocPicture pic = para.AppendPicture(Image.FromFile(@"E:\Ankita\Project\ConversionFromHTMLtoDOC\DocWithHeader\GCHDA.png"));
            //pic.Height = 22;
            //pic.Width = 30;
            //pic.TextWrappingStyle = TextWrappingStyle.;

            ////Header Text
            //TextRange tr = para.AppendText("Microsoft Technology");
            //tr.CharacterFormat.FontName = "Impact";
            //tr.CharacterFormat.FontSize = 12;
            //tr.CharacterFormat.TextColor = Color.DarkBlue;
            //tr.CharacterFormat.ClearFormatting();
            
            //Save and Launch
            doc.SaveToFile("E://ImageHeader.docx", FileFormat.Docx);
            doc.Document.Replace("Evaluation Warning: The document was created with Spire.Doc for .NET.", "", false, true);
            System.Diagnostics.Process.Start("ImageHeader.docx");
        }
    }
}