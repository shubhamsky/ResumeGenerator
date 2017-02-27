using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.Rendering;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace pdfDocumentDemo
{
    class Program
    {
        static void Main(string[] args)
        {          
            PdfDocument document = new PdfDocument();
            CreateDocument(document);
           
            // Save the document...
            const string filename = "HelloWorld.pdf";
          //  pdfRenderer.PdfDocument.Save(filename);
            document.Save(filename);
            // ...and start a viewer.
            Process.Start(filename);
        }

        /// <summary>
        /// Creates Document 
        /// </summary>
        static void CreateDocument(PdfDocument leftPanelPdfDoc)
        {
            //Set LeftPanelData and get XGraphics Object
            XGraphics gfx = setLeftPanelData(leftPanelPdfDoc);

            // Create a new MigraDoc document
            Document rightPanelMigraDoc = new Document();
            // Define Style for Resume
            DefineStylesforResume(rightPanelMigraDoc);
            // Add a section to the document
            Section section = rightPanelMigraDoc.AddSection();

            //set page setup
            string logoPath = "c:/users/shubham yadav/documents/visual studio 2015/Projects/pdfDocumentDemo/pdfDocumentDemo/Images/grad_logo_small.png";

            //setup RightPanel Layout
            SetRightpanel(section, logoPath);

            //Set Right panel Data
            SetRightPanelData(rightPanelMigraDoc);

            //Merge Left and Right Panel
            MergePanels(rightPanelMigraDoc, leftPanelPdfDoc, gfx);
           
        }

        /// <summary>
        /// set Data in Left pane i.e PdfSharp document
        /// </summary>
        /// <param name="leftPanelPdfDoc"></param>
        /// <returns> XGraphics Ojbect</returns>
        public static XGraphics setLeftPanelData(PdfDocument leftPanelPdfDoc)
        {
            XFont font = new XFont("Verdana", 20, XFontStyle.BoldItalic);
            PdfPage page = leftPanelPdfDoc.AddPage();
            XGraphics gfx = CreatePageLayout(page);

            #region setDimension Commented
            //setDimension
            // XRect A4Rect = new XRect(0, 0, A4Width, A4Height);
            // XRect rect = GetRect(index);

            //commented code for setting dimension
            //BeginContainer / EndContainer for transformation
            //XGraphicsContainer container = gfx.BeginContainer(rect, A4Rect, XGraphicsUnit.Point);
            //gfx.EndContainer(container);
            #endregion

            //leftPanelMigraDoc
            Document leftPanelMigraDoc = new Document();

            // Define Style for Left Panel
            //DefineStylesforResume(rightPanelMigraDoc);

            // Add a section to the document
            Section section = leftPanelMigraDoc.AddSection();

            //set page setup
            SetLeftpanel(section);

            //Set Left panel Data
            SetLeftPanelData(leftPanelMigraDoc);

            //Merge Left Panel with pdf Document
            MergePanels(leftPanelMigraDoc, leftPanelPdfDoc, gfx);
            
            return gfx;
        }


        /// <summary>
        /// set Data in Right Panel i.e. MigraDoc document
        /// </summary>
        /// <param name="document"></param>
        public static void SetRightPanelData(Document rightPanelMigraDoc)
        {
            #region RightpanelData
            string scoreImg = "";
            var hrBorder = new Border();
            hrBorder.Width = "1pt";
            for (int i = 1; i < 5; i++)
            {
                string score = i.ToString();

                scoreImg = "c:/users/shubham yadav/documents/visual studio 2015/Projects/pdfDocumentDemo/pdfDocumentDemo/Images/score_" + score + ".png";

                Image scoreImg1 = rightPanelMigraDoc.LastSection.AddImage(scoreImg);
                scoreImg1.Height = "1.5cm";
                scoreImg1.Width = "1.5cm";
                scoreImg1.LockAspectRatio = true;
                scoreImg1.Left = ShapePosition.Left;
                scoreImg1.RelativeVertical = RelativeVertical.Line;
                scoreImg1.RelativeHorizontal = RelativeHorizontal.Margin;
                scoreImg1.WrapFormat.Style = WrapStyle.Through;

                Paragraph occDetailsHead = rightPanelMigraDoc.LastSection.AddParagraph();
                occDetailsHead.AddText("This Is Header Name : ");
                occDetailsHead.Format.LeftIndent = 50;
                occDetailsHead.Format.SpaceBefore = 12;

                string temp = "Loboreet autpat, quis adigna conse dipit la consed exeril et utpatetuer autat, voloboreet,consequamet ilit nos aut in henit ullam, sim doloreratis dolobore tat, venim quissequat. Nisci tatlaor ametumsan vulla feuisim ing eliquisi tatum autat, velenisit iustionsed tis dunt exerostrud dolore verae.";
                Paragraph occDetailsDesc = rightPanelMigraDoc.LastSection.AddParagraph(temp, "Heading3");

                occDetailsDesc.Format.LeftIndent = 50;
                occDetailsDesc.AddLineBreak();
                occDetailsDesc.Format.Borders.Bottom = hrBorder.Clone();

                occDetailsDesc.Format.LineSpacing = 0;
                occDetailsDesc.Format.SpaceBefore = 18;
                occDetailsDesc.Format.SpaceAfter = 18;
                occDetailsDesc.AddLineBreak();

            }

            for (int i = 1; i < 5; i++)
            {
                string score = i.ToString();

                scoreImg = "c:/users/shubham yadav/documents/visual studio 2015/Projects/pdfDocumentDemo/pdfDocumentDemo/Images/score_" + score + ".png";

                Image scoreImg1 = rightPanelMigraDoc.LastSection.AddImage(scoreImg);
                scoreImg1.Height = "1.5cm";
                scoreImg1.Width = "1.5cm";
                scoreImg1.LockAspectRatio = true;
                scoreImg1.Left = ShapePosition.Left;
                scoreImg1.RelativeVertical = RelativeVertical.Line;
                scoreImg1.RelativeHorizontal = RelativeHorizontal.Margin;
                scoreImg1.Top = 10;
                scoreImg1.WrapFormat.Style = WrapStyle.Through;

                Paragraph occDetailsHead = rightPanelMigraDoc.LastSection.AddParagraph();
                occDetailsHead.AddText("This Is Header Name : ");
                occDetailsHead.Format.LeftIndent = 50;
                occDetailsHead.Format.SpaceBefore = 12;

                string temp = "Loboreet autpat, quis adigna conse dipit la consed exeril et utpatetuer autat, voloboreet,consequamet ilit nos aut in henit ullam, sim doloreratis dolobore tat, venim quissequat. Nisci tatlaor ametumsan vulla feuisim ing eliquisi tatum autat, velenisit iustionsed tis dunt exerostrud dolore verae.";
                Paragraph occDetailsDesc = rightPanelMigraDoc.LastSection.AddParagraph(temp, "Heading3");

                occDetailsDesc.Format.LeftIndent = 50;
                occDetailsDesc.AddLineBreak();
                occDetailsDesc.Format.Borders.Bottom = hrBorder.Clone();

                occDetailsDesc.Format.LineSpacing = 0;
                occDetailsDesc.Format.SpaceBefore = 18;
                occDetailsDesc.Format.SpaceAfter = 18;
                occDetailsDesc.AddLineBreak();

            }
            for (int i = 1; i < 5; i++)
            {
                string score = i.ToString();

                scoreImg = "c:/users/shubham yadav/documents/visual studio 2015/Projects/pdfDocumentDemo/pdfDocumentDemo/Images/score_" + score + ".png";

                Image scoreImg1 = rightPanelMigraDoc.LastSection.AddImage(scoreImg);
                scoreImg1.Height = "1.5cm";
                scoreImg1.Width = "1.5cm";
                scoreImg1.LockAspectRatio = true;
                scoreImg1.Left = ShapePosition.Left;
                scoreImg1.RelativeVertical = RelativeVertical.Line;
                scoreImg1.RelativeHorizontal = RelativeHorizontal.Margin;
                scoreImg1.Top = 10;
                scoreImg1.WrapFormat.Style = WrapStyle.Through;

                Paragraph occDetailsHead = rightPanelMigraDoc.LastSection.AddParagraph();
                occDetailsHead.AddText("This Is Header Name : ");
                occDetailsHead.Format.LeftIndent = 50;
                occDetailsHead.Format.SpaceBefore = 12;

                string temp = "Loboreet autpat, quis adigna conse dipit la consed exeril et utpatetuer autat, voloboreet,consequamet ilit nos aut in henit ullam, sim doloreratis dolobore tat, venim quissequat. Nisci tatlaor ametumsan vulla feuisim ing eliquisi tatum autat, velenisit iustionsed tis dunt exerostrud dolore verae.";
                Paragraph occDetailsDesc = rightPanelMigraDoc.LastSection.AddParagraph(temp, "Heading3");

                occDetailsDesc.Format.LeftIndent = 50;
                occDetailsDesc.AddLineBreak();
                occDetailsDesc.Format.Borders.Bottom = hrBorder.Clone();

                occDetailsDesc.Format.LineSpacing = 0;
                occDetailsDesc.Format.SpaceBefore = 18;
                occDetailsDesc.Format.SpaceAfter = 18;
                occDetailsDesc.AddLineBreak();

            }
            #endregion
        }

        /// <summary>
        /// Setup RightPanel Page size, header and Footer
        /// </summary>
        /// <param name="section"></param>
        /// <param name="logoPath"></param>
        public static void SetRightpanel(Section section, string logoPath)
        {
            #region pagesetup
            section.PageSetup.DifferentFirstPageHeaderFooter = true;
            section.PageSetup.LeftMargin = "8cm";
            section.PageSetup.RightMargin = "1cm";
            section.PageSetup.TopMargin = "1cm";
            section.PageSetup.BottomMargin = Unit.FromCentimeter(2.5);
            section.PageSetup.StartingNumber = 1;
            #endregion

            #region Footer
            //Logo           
            Image imageLogo = section.Footers.FirstPage.AddImage(logoPath);
            imageLogo.Height = "1cm";
            imageLogo.LockAspectRatio = true;
            imageLogo.RelativeVertical = RelativeVertical.Line;
            imageLogo.RelativeHorizontal = RelativeHorizontal.Margin;
            imageLogo.Top = ShapePosition.Bottom;
            imageLogo.Left = ShapePosition.Right;
            imageLogo.WrapFormat.Style = WrapStyle.Through;

            //page number on First Page
            Paragraph pagenumber = section.Footers.FirstPage.AddParagraph();
            pagenumber.AddText("Page ");
            pagenumber.Format.LineSpacing = 0;
            pagenumber.Format.SpaceBefore = 15;
            pagenumber.Format.SpaceAfter = 5;
            pagenumber.AddTab();
            pagenumber.AddPageField();
            pagenumber.Format.Alignment = ParagraphAlignment.Center;

            //border
            var hrBorder = new Border();
            hrBorder.Width = "1pt";
            hrBorder.Color = Colors.Gray;

            section.Footers.Primary.Add(imageLogo.Clone());

            //pagenumber on all footer
            section.Footers.Primary.Add(pagenumber.Clone());        
            #endregion
        }

        //-- Method for Coloring the panel in page
        //-- Returns XGraphics Object gfx
        public static  XGraphics CreatePageLayout(PdfPage page)
        {
            //Create Graphics object from page
            XGraphics gfx = XGraphics.FromPdfPage(page);
            gfx.MUH = PdfFontEncoding.Unicode;
            gfx.MFEH = PdfFontEmbedding.Default;
            //---Color rectangle dimension--------------------------------------------------------------------
            XUnit horizontalOffset = XUnit.FromCentimeter(0);
            XUnit verticalOffset = XUnit.FromCentimeter(0);
            XUnit columnWidth = XUnit.FromCentimeter(7).Point;
            XUnit columnHeight = page.Height;
            XRect columnRect = new XRect(horizontalOffset, verticalOffset, columnWidth, columnHeight);
            //---Color on page--------------------------------------------------------------------------
            gfx.DrawRectangle(new XSolidBrush(XColor.FromArgb(98, 183, 229)), columnRect);
            return gfx;
        }

        static void DefineStylesforResume(Document document)
        {
            Style style = document.Styles["Normal"];
            style.Font.Name = "Montserrat";
            style = document.Styles["Heading1"];
            style.Font.Name = "Oswald";
            style.Font.Size = 22;
            style.Font.Bold = true;
            style.Font.Color = Colors.SaddleBrown;
            style.ParagraphFormat.PageBreakBefore = false;
            style.ParagraphFormat.SpaceAfter = 6;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            style = document.Styles["Heading2"];
            style.Font.Size = 12;
            style.Font.Name = "Montserrat";
            style.Font.Bold = true;
            style.Font.Color = Colors.Black;
            style.ParagraphFormat.PageBreakBefore = false;
            style.ParagraphFormat.SpaceBefore = 3;
            style.ParagraphFormat.SpaceAfter = 10;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            style = document.Styles["Heading3"];
            style.Font.Size = 11;
            style.Font.Name = "Montserrat";
            style.Font.Bold = false;
            style.Font.Color = Colors.Gray;
            style.ParagraphFormat.PageBreakBefore = false;
            style.ParagraphFormat.SpaceBefore = 0;
            style.ParagraphFormat.SpaceAfter = 10;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        }

        static XRect GetRect(int index)
        {
            XRect rect = new XRect(0, 0, A4Width / 2 * 0.9, A4Height / 2 * 0.9);
            rect.X = (index % 2) * A4Width / 2 + A4Width * 0.05 / 2;
            rect.Y = (index / 2) * A4Height / 2 + A4Height * 0.05 / 2;
            return rect;
        }

        //page Dimension constants
        static double A4Width = XUnit.FromCentimeter(21).Point;
        static double A4Height = XUnit.FromCentimeter(29.7).Point;

        /// <summary>
        /// Merge Left and Right Panels
        /// </summary>
        /// <param name="rightPanelMigraDoc"></param>
        /// <param name="leftPanelPdfDoc"></param>
        /// <param name="gfx"></param>
        public static void MergePanels(Document migraDocument, PdfDocument leftPanelPdfDoc, XGraphics gfx)
        {
            PdfPage page = null;

            //concatenation of Left and Right Panels Core of Resume
            // Create a renderer and prepare (layout) the Migradoc document
            MigraDoc.Rendering.DocumentRenderer docRenderer = new DocumentRenderer(migraDocument);
            docRenderer.PrepareDocument();

            //get pagecount of the rightpanelMigraDoc
            int rightPageCount = docRenderer.FormattedDocument.PageCount;
            int leftPageCount = leftPanelPdfDoc.PageCount;
            // Loop to add leftPanel pages equal to rightPanel and Merge by rendering
            int index = 0;
            do
            {
                //for single left panel page only.
                if (index == 0)
                {
                    docRenderer.RenderPage(gfx, index + 1);
                    index++;
                }
                else
                {
                    page = leftPanelPdfDoc.AddPage();
                    gfx = CreatePageLayout(page);
                    // Render the page. Note that page numbers start with 1.
                    docRenderer.RenderPage(gfx, index + 1);
                    index++;
                }
            }
            while (index < rightPageCount);
        }

        public static void SetLeftpanel(Section section)
        {
            #region pagesetup
            section.PageSetup.DifferentFirstPageHeaderFooter = true;
            section.PageSetup.LeftMargin = "0.5cm";
            section.PageSetup.RightMargin = "14.5cm";
            section.PageSetup.TopMargin = "0.5cm";
            section.PageSetup.BottomMargin = Unit.FromCentimeter(0.5);
            //page number on First Page
            #endregion
        }

        /// <summary>
        /// set Data in Right Panel i.e. MigraDoc document
        /// </summary>
        /// <param name="document"></param>
        public static void SetLeftPanelData(Document leftPanelMigraDoc)
        {
            #region LeftPanelData
            //set layout of leftpanel Migradoc
            DefineStylesforLeftPanel(leftPanelMigraDoc);
            string scoreImg = "";
            var hrBorder = new Border();
            hrBorder.Width = "1pt";
            hrBorder.Color = Colors.White;
            for (int i = 1; i < 5; i++)
            {
                string score = i.ToString();

                scoreImg = "c:/users/shubham yadav/documents/visual studio 2015/Projects/pdfDocumentDemo/pdfDocumentDemo/Images/score_" + score + ".png";

                Image scoreImg1 = leftPanelMigraDoc.LastSection.AddImage(scoreImg);
                scoreImg1.Height = "0.8cm";
                scoreImg1.Width = "0.8cm";
                scoreImg1.LockAspectRatio = true;
                scoreImg1.Left = ShapePosition.Left;
                scoreImg1.RelativeVertical = RelativeVertical.Line;
                scoreImg1.RelativeHorizontal = RelativeHorizontal.Margin;
                scoreImg1.WrapFormat.Style = WrapStyle.Through;

                Paragraph occDetailsHead = leftPanelMigraDoc.LastSection.AddParagraph();
                occDetailsHead.AddFormattedText("This Is Header Name : ", "Heading1");
                //occDetailsHead.AddFormattedText("This Is Header Name : ", "Heading2");
               occDetailsHead.Format.LeftIndent = 25;
               // occDetailsHead.Format.SpaceBefore = 12;

                string temp = "Loboreet autpat, quis adigna conse dipit la consed exeril et utpatetuer autat, voloboreet,consequamet ilit nos aut in henit ullam, sim doloreratis dolobore tat, venim quissequat. Nisci tatlaor ametumsan vulla feuisim ing eliquisi tatum autat, velenisit iustionsed tis dunt exerostrud dolore verae.";
                Paragraph occDetailsDesc = leftPanelMigraDoc.LastSection.AddParagraph(temp);

              //  occDetailsDesc.Format.LeftIndent = 25;
                occDetailsDesc.AddLineBreak();
                occDetailsDesc.Format.Borders.Bottom = hrBorder.Clone();
                occDetailsDesc.Style = "Heading3";

                occDetailsDesc.Format.LineSpacing = 0;
                occDetailsDesc.Format.SpaceBefore = 10;
                occDetailsDesc.Format.SpaceAfter = 15;
                occDetailsDesc.AddLineBreak();

            }
            for (int i = 1; i < 5; i++)
            {
                string score = i.ToString();

                scoreImg = "c:/users/shubham yadav/documents/visual studio 2015/Projects/pdfDocumentDemo/pdfDocumentDemo/Images/score_" + score + ".png";

                Image scoreImg1 = leftPanelMigraDoc.LastSection.AddImage(scoreImg);
                scoreImg1.Height = "0.8cm";
                scoreImg1.Width = "0.8cm";
                scoreImg1.LockAspectRatio = true;
                scoreImg1.Left = ShapePosition.Left;
                scoreImg1.RelativeVertical = RelativeVertical.Line;
                scoreImg1.RelativeHorizontal = RelativeHorizontal.Margin;
                scoreImg1.WrapFormat.Style = WrapStyle.Through;

                Paragraph occDetailsHead = leftPanelMigraDoc.LastSection.AddParagraph();
                occDetailsHead.AddFormattedText("This Is Header Name : ", "Heading1");
                //occDetailsHead.AddFormattedText("This Is Header Name : ", "Heading2");
                occDetailsHead.Format.LeftIndent = 25;
                // occDetailsHead.Format.SpaceBefore = 12;

                string temp = "Loboreet autpat, quis adigna conse dipit la consed exeril et utpatetuer autat, voloboreet,consequamet ilit nos aut in henit ullam, sim doloreratis dolobore tat, venim quissequat. Nisci tatlaor ametumsan vulla feuisim ing eliquisi tatum autat, velenisit iustionsed tis dunt exerostrud dolore verae.";
                Paragraph occDetailsDesc = leftPanelMigraDoc.LastSection.AddParagraph(temp);

                //  occDetailsDesc.Format.LeftIndent = 25;
                occDetailsDesc.AddLineBreak();
                occDetailsDesc.Format.Borders.Bottom = hrBorder.Clone();
                occDetailsDesc.Style = "Heading3";

                occDetailsDesc.Format.LineSpacing = 0;
                occDetailsDesc.Format.SpaceBefore = 10;
                occDetailsDesc.Format.SpaceAfter = 15;
                occDetailsDesc.AddLineBreak();

            }
            #endregion
        }

        static void DefineStylesforLeftPanel(Document document)
        {
            Style style = document.Styles["Normal"];
            style.Font.Name = "Montserrat";
            style = document.Styles["Heading1"];
            style.Font.Name = "Montserrat";
            style.Font.Size = 12;
            style.Font.Bold = true;
            style.Font.Color = Colors.White;
            style.ParagraphFormat.PageBreakBefore = false;
            style.ParagraphFormat.LeftIndent = 30;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Center;

            style = document.Styles["Heading2"];
            style.Font.Size = 10;
            style.Font.Name = "Montserrat";
            style.Font.Bold = true;
            style.Font.Color = Colors.White;
            style.ParagraphFormat.PageBreakBefore = false;
            //style.ParagraphFormat.SpaceBefore = 3;
            //style.ParagraphFormat.SpaceAfter = 10;
            style.ParagraphFormat.LeftIndent = 25;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Left;

            style = document.Styles["Heading3"];
            style.Font.Size = 10;
            style.Font.Name = "Montserrat";
            style.Font.Bold = false;
            style.Font.Color = Colors.White;
            style.ParagraphFormat.PageBreakBefore = false;
            style.ParagraphFormat.SpaceBefore = 10;
            style.ParagraphFormat.SpaceAfter = 15;
            style.ParagraphFormat.LeftIndent = 25;
            style.ParagraphFormat.Alignment = ParagraphAlignment.Left;
        }
    }
}
