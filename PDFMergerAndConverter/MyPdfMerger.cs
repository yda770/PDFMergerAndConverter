using iText.IO.Font;
using iText.Kernel.Font;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Utils;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Layout;
using iText.Layout.Renderer;
using iText.Kernel.Colors;
using iText.Kernel.Geom;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using iText.Kernel.Pdf.Colorspace;
using iText.Layout.Properties;
using iText.Layout.Borders;

namespace PDFMergerAndConverter
{   
    class MyPdfMerger
    {
        private PdfDocument outPdf;
        private PdfMerger merger;
        private int pageBrakes;
        const string TitleA = "Document name";
        const string TitleB = "Pages";

        public MyPdfMerger(string outputPDFpath)
        {
            this.outPdf = new PdfDocument(new PdfWriter(outputPDFpath));
            this.merger = new PdfMerger(this.outPdf);
            this.pageBrakes = 0;
        }

        public string addPdf(string source)
        {
            if (!File.Exists(source))
            {
                var pdfSorce = new PdfDocument(new PdfReader(source.Trim()));
                merger.Merge(pdfSorce, 1, pdfSorce.GetNumberOfPages());

                Encoding defaultEncoding = Encoding.Default;
                byte[] bytes = defaultEncoding.GetBytes(TitleA);
                Encoding encoding2 = Encoding.UTF8;
                string TitleA_CNV = encoding2.GetString(bytes);

                defaultEncoding = Encoding.Default;
                bytes = defaultEncoding.GetBytes(TitleB);
                encoding2 = Encoding.UTF8;
                string TitleB_CNV = encoding2.GetString(bytes);

                this.AddPageTitle(this.outPdf, 
                                  this.outPdf.GetNumberOfPages() - pdfSorce.GetNumberOfPages() + 1,
                                  TitleA + " " + source.Trim() + " (" + pdfSorce.GetNumberOfPages() + " " + TitleB + ")");
                pdfSorce.Close();
            }
            return "";
        }

        private void AddPageTitle(PdfDocument pdfDoc, int pageNumber, string pageTitle)
        { 
            PdfPage page = pdfDoc.GetPage(pageNumber);
            
            Paragraph paragraph = new Paragraph(pageTitle).
                  SetMargin(3).
                  SetMultipliedLeading(1).
                  SetFont(PdfFontFactory.CreateFont(FontConstants.TIMES_ROMAN));

            float fontSizeF = 20;
 
            //paragraph.SetBorder(new SolidBorder(1));
            float allowedWidth = 185;

            Canvas canvas = new Canvas(new PdfCanvas(page), pdfDoc, page.GetMediaBox());
            RootRenderer canvasRenderer = canvas.GetRenderer();
            paragraph.SetFontSize(fontSizeF);
            paragraph.SetBackgroundColor(ColorConstants.GRAY, 0.7f); // Just to see that text is aligned correctly
            paragraph.SetFontColor(ColorConstants.WHITE);
            paragraph.SetWidth(page.GetPageSize().GetWidth() - 30);

            while (paragraph.CreateRendererSubTree().SetParent(canvasRenderer).Layout(new LayoutContext(new LayoutArea(pageNumber, new iText.Kernel.Geom.Rectangle(allowedWidth, fontSizeF * 3)))).GetStatus() != LayoutResult.FULL)
            {
                paragraph.SetFontSize(fontSizeF - 5);
            }

            float xCoord = page.GetPageSize().GetWidth() / 2;
            float yCoord = page.GetPageSize().GetHeight() -30;

            
            canvas.ShowTextAligned(paragraph, xCoord, yCoord, TextAlignment.CENTER);
            canvas.Close();
            //pdfDocument.close();
        }
        public void ClosePdf()
        {
            this.outPdf.Close();
        }

        public bool CheckMimeType()
        {
            return false;
        }

        public void AddImageToPdf(string imageSource)
        {
            Document document = new Document(this.outPdf);

            iText.IO.Image.ImageData imageData = iText.IO.Image.ImageDataFactory.Create(imageSource.Trim());
            Image image = new Image(imageData);
            image.SetWidth(this.outPdf.GetDefaultPageSize().GetWidth() - 50);
            image.SetAutoScaleHeight(true);

            int pagesNum = this.outPdf.GetNumberOfPages();
            int currentPageBreaks = this.pageBrakes;
                
            for (int pageNum = 0; pageNum < pagesNum - currentPageBreaks; pageNum++)
            {
                document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                this.pageBrakes++;
            }

            if (pageBrakes > 0)
            {
                this.pageBrakes--;
            }
            
            document.Add(image);

            this.AddPageTitle(this.outPdf,
                  this.outPdf.GetNumberOfPages(),
                  TitleA + " " + imageSource.Trim() + " (1 " + TitleB + ")");

        }
        public void AddTextToPdf(string textSource)
        {
            Document document = new Document(this.outPdf);

            iText.IO.Image.ImageData imageData = iText.IO.Image.ImageDataFactory.Create(imageSource.Trim());
            Image image = new Image(imageData);
            image.SetWidth(this.outPdf.GetDefaultPageSize().GetWidth() - 50);
            image.SetAutoScaleHeight(true);

            int pagesNum = this.outPdf.GetNumberOfPages();
            int currentPageBreaks = this.pageBrakes;

            for (int pageNum = 0; pageNum < pagesNum - currentPageBreaks; pageNum++)
            {
                document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                this.pageBrakes++;
            }

            if (pageBrakes > 0)
            {
                this.pageBrakes--;
            }

            document.Add(image);

            this.AddPageTitle(this.outPdf,
                  this.outPdf.GetNumberOfPages(),
                  TitleA + " " + imageSource.Trim() + " (1 " + TitleB + ")");

        }
    }
}
