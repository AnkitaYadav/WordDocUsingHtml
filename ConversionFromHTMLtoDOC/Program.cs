using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NotesFor.HtmlToOpenXml;
using System.Reflection;
using System.Resources;
using System.Globalization;
using System.Collections.Generic;
using System;
using System.Net;
using System.Web.Script.Serialization;
using System.Text;

namespace ConversionFromHTMLtoDOC
{
    class Program
    {
        static void Main(string[] args)
        {
            WebClient client = new WebClient();
            var meetingResponse = client.DownloadString("http://38.118.71.177/nyacuradealers/service//meeting/meetingDetails?meetingId=12&associationId=1");
            var attendeesResponse = client.DownloadString("http://38.118.71.177/nyacuradealers/service/meeting/attendees?meetingId=12");

            JavaScriptSerializer serialser = new JavaScriptSerializer();
            var meetingInfo = serialser.Deserialize<MeetingModel>(meetingResponse);
            var attendeesInfo = serialser.Deserialize<IEnumerable<Attendee>>(attendeesResponse);
            const string filename = @"E:\Ankita\MeetingWorddoc/testWordDoc.docx";

            //string html = File.ReadAllText(@"E:\Ankita\Project\ConversionFromHTMLtoDOC\ConversionFromHTMLtoDOC\test2.html");
            StringBuilder meeting = new StringBuilder();
            //var headHtml = @"<!DOCTYPE HTML PUBLIC><html  xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><title></title><style type=""text/css"">.light-dark-bg{display:block;background-color:lightGray;width:100%;},body{background:#f3f3f3}*,.page{box-sizing:border-box}*{margin:0;padding:0;font-family:arial;font-size:13px}.clearfix:after{content:'';visibility:hidden;display:block;height:0;clear:both}.fleft{float:left}.fright{float:right}.logo img{width:230px;position:relative;left:-15px}.sec-head{text-transform:uppercase;font-size:15px;color:#07476c;padding:10px 0 5px}.width-100px{width:100px;display:inline-block}.sub-sec{font-size:13px;text-transform:none}.p-0{padding:0}.p-5{padding:5px}.p-0-5{padding-left:5px;padding-right:5px}.p-5-0{padding-top:5px;padding-bottom:5px}.b{font-weight:700}.u{text-decoration:underline}.page{max-width:1000px;margin:auto;padding:40px;background:#fff}.no-list-style{list-style:none}.list-style-decimal{list-style-type:decimal}.list-style-bullit{list-style-type:circle}.address{font-size:12px}.body-section{margin:20px 0;padding-bottom:10px}.attend-list li{width:50%;float:left;font-size:13px}.m-10-0{margin-top:10px;margin-bottom:10px}.m-l-15{margin-left:15px}.m-l-25{margin-left:25px}.m-l-30{margin-left:30px}.m-l-50{margin-left:50px};@page{size:21cm 29.7cmt; margin:1cm 1cm 1cm 1cm; mso-page-orientation: portrait;}; </style></head><body><section class='body'>";
            var headHtml = @"<!DOCTYPE HTML PUBLIC><html  xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'><head><title></title>
          <style type=""text/css"">
       
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
            font-family: arial;
            font-size: 13px;
        }

        body {
            background: #f3f3f3;
        }


        .clearfix:after {
            content: ' ';
            visibility: hidden;
            display: block;
            height: 0;
            clear: both;
        }

        .fleft {
            float: left;
        }

        .fright {
            float: right;
        }

        .logo img {
            width: 230px;
            position: relative;
            left: -15px;
        }

        .sec-head {
            text-transform: uppercase;
            font-size: 15px;
            color: #07476c;
            padding: 10px 0;
            padding-bottom: 20px;
            margin-bottom:5px;
        }

        .light-dark-bg {
            background-color: lightgray;
            width: 100%;
        }

        .width-100px {
            width: 100px;
            display: inline-block;
        }

        .sub-sec {
            font-size: 15px;
            text-transform: none;
        }

        .p-0 {
            padding: 0;
        }

        .p-5 {
            padding: 5px;
        }

        .p-0-5 {
            padding-left: 5px;
            padding-right: 5px;
        }

        .p-5-0 {
            padding-top: 5px;
            padding-bottom: 5px;
        }

        .b {
            font-weight: bold;
        }

        .u {
            text-decoration: underline;
        }

        .page {
            max-width: 1000px;
            margin: auto;
            box-sizing: border-box;
            padding: 40px;
            background: #fff;
        }

        .no-list-style {
            list-style: none;
        }

        .list-style-decimal {
            list-style-type: decimal;
        }

        .list-style-bullit {
            list-style-type: circle;
        }

        .address {
            font-size: 12px;
        }

        .body-section {
            margin: 0px 0;
            padding-bottom: 0px;
        }

        .attend-list li {
            width: 50%;
            float: left;
            font-size: 13px;
        }

        .m-10-0 {
            margin-top: 5px;
            margin-bottom: 0px;
        }

        .m-l-15 {
            margin-left: 15px;
        }
        .m-l-5
        {
          margin-left: 5px;
        }

        .m-l-25 {
            margin-left: 25px;
        }

        .m-l-30 {
            margin-left: 30px;
            margin-top: 5px;
            margin-bottom: 5px;
            
        }

        .m-l-50 {
            margin-left: 50px;
        }
        .p-20-0 {
            padding-top: 20px;
            padding-bottom: 20px;
        }
        </style>
         </head>
          <body>
          <section class='body'>";

            meeting.Append(headHtml);
            string attendeesHtml = @"<section class='body-section clearfix'><p class='b sec-head'>ATTENDEES<div class=attend-list><ul class='list-style-decimal m-l-30'>[ATTENDEE]</ul></div></section>";
            string attendee = string.Empty;
            foreach (var atnd in attendeesInfo)
            {
                attendee = attendee + "<li>" + atnd.Name + ", " + atnd.Company + "</li>";
            }
            meeting.Append(attendeesHtml.Replace("[ATTENDEE]", attendee));

            string groupDiscussionHtml = @"<section class='body-section clearfix light-dark-bg'><p class='b sec-head'>MASTER STRATEGY, CREATIVE, GROUP DECISION<div class='content m-l-50'>[GROUPDECISION]</div></section>";
            meeting.Append(groupDiscussionHtml.Replace("[GROUPDECISION]", meetingInfo.GroupDecision));
            meeting.Append("<p class='sec-head b'>MEETING SUMMARY</p>");
            foreach (var meetingAction in meetingInfo.MeetingActions)
            {
                var meetingHtml = File.ReadAllText(@"E:\Ankita\Project\ConversionFromHTMLtoDOC\ConversionFromHTMLtoDOC\test2.html");

                StringBuilder meetingActionHtm = new StringBuilder(meetingHtml);
                meetingActionHtm.Replace("[MeetingAction]", meetingAction.ActionName).Replace("[Discussion]", meetingAction.Discussion)
                               .Replace("[Conclusions]", meetingAction.Conclusion);
                string actionItems = string.Empty;
                foreach (var actionItem in meetingAction.ActionItems)
                {
                    actionItems = actionItems + "<li>" + actionItem.Text + "</li>";
                }
                meetingActionHtm.Replace("[ActionItem]", actionItems);
                meeting.Append(meetingActionHtm);
            }

            string performancehtml = @"<section class='body-section clearfix'>
                <p class='sec-head b'>PERFORMANCE RECAP</p>
                <p class='m-l-50'>[PERFORMANCEPOINT]</p>
                <div class='content m-l-50'>
                    <p class='sec-head sub-sec p-0 b'>Retail Sales Performance</p>
                    <p class='b u'> YDT </p>
                    <div><span >DAA</span> <span>        [DAA]%</span></div> 
                     <div><span >District</span> <span>      [DISTRICT]%</span></div> 
                     <div><span  >Zone</span> <span>      [ZONE]%</span></div> 
                     <div><span  >Nation</span> <span>        [NATION]%</span></div> 
                   
                </div>
            </section>";
            meeting.Append(performancehtml.Replace("[PERFORMANCEPOINT]", meetingInfo.PerformancePoints)
                          .Replace("[DAA]", meetingInfo.Daa)
                           .Replace("[DISTRICT]", meetingInfo.District)
                            .Replace("[NATION]", meetingInfo.Nation)
                             .Replace("[ZONE]", meetingInfo.Zone));
            var actionItemSummaryHml = @"<section class='body-section clearfix '>
                                           <p class='sec-head b'> ACTION ITEMS </p>
                                                [ActionItemSummary]
                                         </section>";
            var summary = "";

            foreach (var meetingAction in meetingInfo.MeetingActions)
            {

                var itemSummary = @"<div class='content m-10-0'><p class=b>[ACTIONTYPE] Action Items<ul class='list-style-decimal m-l-30'>[ACTIONITEM]</ul></div>";
                StringBuilder meetingActionSummaryHtm = new StringBuilder(itemSummary);
                meetingActionSummaryHtm.Replace("[ACTIONTYPE]", meetingAction.ActionName);
                string actionItems = string.Empty;
                foreach (var actionItem in meetingAction.ActionItems)
                {
                    actionItems = actionItems + "<li>" + actionItem.Text + "</li>";
                }
                meetingActionSummaryHtm.Replace("[ACTIONITEM]", actionItems);
                // meetingActionSummaryHtm.Append(actionItems);
                summary += Convert.ToString(meetingActionSummaryHtm);
            }
            meeting.Append(actionItemSummaryHml.Replace("[ActionItemSummary]", summary));


            var bottomHtml = @"</section></body></html> ";
            meeting.Append(bottomHtml);
            var result = PreMailer.Net.PreMailer.MoveCssInline(meeting.ToString());
            string html = result.Html;
            try
            {
                if (File.Exists(filename)) File.Delete(filename);

                using (MemoryStream generatedDocument = new MemoryStream())
                {
                    using (WordprocessingDocument package = WordprocessingDocument.Create(generatedDocument, WordprocessingDocumentType.Document))
                    {
                        MainDocumentPart mainPart = package.MainDocumentPart;
                        if (mainPart == null)
                        {
                            mainPart = package.AddMainDocumentPart();
                            new Document(new Body()).Save(mainPart);
                        }

                        HtmlConverter converter = new HtmlConverter(mainPart);
                        converter.ParseHtml(html);

                        mainPart.Document.Save();
                    }

                    File.WriteAllBytes(filename, generatedDocument.ToArray());
                }

               



            }
            catch (System.Exception)
            {

                throw;
            }



            //ChangeHeader(filename);

            System.Diagnostics.Process.Start(filename);
        }
        public static void AddHeader(string wordFile, string header)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordFile, true))
            {
                foreach (HeaderPart hp in wordDoc.MainDocumentPart.HeaderParts)
                {




                }
            }
        }

        static void ChangeHeader(String documentPath)
        {
            // Replace header in target document with header of source document.
            using (WordprocessingDocument document = WordprocessingDocument.Open(documentPath, true))
            {
                // Get the main document part
                MainDocumentPart mainDocumentPart = document.MainDocumentPart;

                // Delete the existing header and footer parts
                mainDocumentPart.DeleteParts(mainDocumentPart.HeaderParts);
                mainDocumentPart.DeleteParts(mainDocumentPart.FooterParts);

                // Create a new header and footer part
                HeaderPart headerPart = mainDocumentPart.AddNewPart<HeaderPart>();
                FooterPart footerPart = mainDocumentPart.AddNewPart<FooterPart>();

                // Get Id of the headerPart and footer parts
                string headerPartId = mainDocumentPart.GetIdOfPart(headerPart);
                string footerPartId = mainDocumentPart.GetIdOfPart(footerPart);

                GenerateHeaderPartContent(headerPart);

                GenerateFooterPartContent(footerPart);

                // Get SectionProperties and Replace HeaderReference and FooterRefernce with new Id
                IEnumerable<SectionProperties> sections = mainDocumentPart.Document.Body.Elements<SectionProperties>();

                foreach (var section in sections)
                {
                    // Delete existing references to headers and footers
                    section.RemoveAllChildren<HeaderReference>();
                    section.RemoveAllChildren<FooterReference>();

                    // Create the new header and footer reference node
                    section.PrependChild<HeaderReference>(new HeaderReference() { Id = headerPartId });
                    section.PrependChild<FooterReference>(new FooterReference() { Id = footerPartId });
                }
            }
        }

        static void GenerateHeaderPartContent(HeaderPart part)
        {
            Header header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            header1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00164C17", RsidRunAdditionDefault = "00164C17" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run();
            Text text1 = new Text();
            text1.Text = "Header";

            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            header1.Append(paragraph1);

            part.Header = header1;
        }

        static void GenerateFooterPartContent(FooterPart part)
        {
            Footer footer1 = new Footer() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            footer1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            footer1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            footer1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            footer1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            footer1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            footer1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            footer1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            footer1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            footer1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            footer1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            footer1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            footer1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            footer1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            footer1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            footer1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00164C17", RsidRunAdditionDefault = "00164C17" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Footer" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run();
            Text text1 = new Text();
            text1.Text = "Footer";

            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            footer1.Append(paragraph1);

            part.Footer = footer1;
        }
    }
}






