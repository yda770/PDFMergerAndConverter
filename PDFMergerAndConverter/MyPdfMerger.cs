using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Utils;
using iText.Layout;
using iText.Layout.Borders;
using iText.Layout.Element;
using iText.Layout.Font;
using iText.Layout.Properties;
using System;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Text;

namespace PDFMergerAndConverter
{
    class MyPdfMerger
    {
        private PdfDocument outPdf;
        private PdfMerger merger;
        private int pageBreaks;
        const string TitleA = "Document name";
        const string TitleB = "Pages";

        public MyPdfMerger(string outputPDFpath)
        {
            this.outPdf = new PdfDocument(new PdfWriter(outputPDFpath));
            this.merger = new PdfMerger(this.outPdf);
            this.pageBreaks = 0;
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
            float fontSizeF = 20;
            float allowedWidth = 185;

            try
            {
                string rootFolder = System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
                PdfPage page = pdfDoc.GetPage(pageNumber);
                
                PdfCanvas pdfCanvas = new PdfCanvas(this.outPdf, pageNumber);
                Rectangle rectangle = new Rectangle(allowedWidth, fontSizeF * 3);
                Canvas canvas = new Canvas(pdfCanvas, rectangle);
                PdfFont font = PdfFontFactory.CreateFont(rootFolder + @"\Arial.ttf", PdfEncodings.IDENTITY_H);
                Text title =
                  new Text(pageTitle).SetFont(font);

                Paragraph paragraph = new Paragraph().Add(title);
                paragraph.SetBackgroundColor(ColorConstants.GRAY, 0.7f);
                paragraph.SetFontColor(ColorConstants.WHITE);
                paragraph.SetWidth(page.GetPageSize().GetWidth() - 30);
                paragraph.SetFontSize(fontSizeF - 7);
                //paragraph.SetBorder(new SolidBorder(1));
                paragraph.SetMargin(3);
                paragraph.SetMultipliedLeading(1);

                float xCoord = page.GetPageSize().GetWidth() / 2;
                float yCoord = page.GetPageSize().GetHeight() - 35;

                canvas.ShowTextAligned(paragraph, xCoord, yCoord, TextAlignment.CENTER);
                canvas.Close();

            }
            catch (System.Exception ex)
            {
                Console.WriteLine(ex.Message);


            }
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
            ImageData imageData = ImageDataFactory.Create(imageSource.Trim());
            Image image = new Image(imageData);
            image.SetWidth(this.outPdf.GetDefaultPageSize().GetWidth() - 50);
            image.SetAutoScaleHeight(true);

            int pagesNum = this.outPdf.GetNumberOfPages();
            int currentPageBreaks = this.pageBreaks;

            this.outPdf.AddNewPage();
            for (int pageNum = 0; pageNum < pagesNum - currentPageBreaks; pageNum++)
            {
                document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                this.pageBreaks++;
            }

            if (pageBreaks > 0)
            {
                this.pageBreaks--;
            }

            document.Add(image);

            this.AddPageTitle(this.outPdf,
                  pagesNum + 1,
                  TitleA + " " + imageSource.Trim() + " (1 " + TitleB + ")");

        }
        public void AddTextToPdf(string textSource)
        {
            Document document = new Document(this.outPdf);

            int NewPages = 0;
            if (!File.Exists(textSource))
            {
                string text = File.ReadAllText(textSource.Trim());

                int pagesNum = this.outPdf.GetNumberOfPages();
                int currentPageBreaks = this.pageBreaks;

                for (int pageNum = 0; pageNum < pagesNum - currentPageBreaks; pageNum++)
                {
                    document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                    this.pageBreaks++;
                    NewPages++;
                }

                if (pageBreaks > 0)
                {
                    this.pageBreaks--;
                }

                string rootFolder = System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
                PdfFont font = PdfFontFactory.CreateFont(rootFolder + @"\Arial.ttf", PdfEncodings.IDENTITY_H);
                document.Add(new Paragraph(text).SetFont(font));

                this.AddPageTitle(this.outPdf,
                        pagesNum + 1,
                        TitleA + " " + textSource.Trim() + " (1 " + TitleB + ")");
            }
        }
        public void AddWordToPdf(string docPath)
        {
            var app = new Microsoft.Office.Interop.Word.Application();
            app.Visible = true;
            int pagesNum = this.outPdf.GetNumberOfPages();
            var doc = app.Documents.Open(docPath.Trim());
            Document document = new Document(this.outPdf, this.outPdf.GetDefaultPageSize(), false);
            doc.ShowGrammaticalErrors = false;
            doc.ShowRevisions = false;
            doc.ShowSpellingErrors = false;

            //Opens the word document and fetch each page and converts to image
            foreach (Microsoft.Office.Interop.Word.Window window in doc.Windows)
            {
                foreach (Microsoft.Office.Interop.Word.Pane pane in window.Panes)
                {
                    for (var i = 1; i <= pane.Pages.Count; i++)
                    {
                        var page = pane.Pages[i];
                        var bits = page.EnhMetaFileBits;
                        
                        try
                        {
                            var ms = new MemoryStream((byte[])(bits));
                            var image = System.Drawing.Image.FromStream(ms);
                            string tempPath = System.IO.Path.GetTempPath();
                            tempPath += @"/image_" + GetTimestamp(new DateTime());
                            var pngTarget = System.IO.Path.ChangeExtension(tempPath, "png");
                            image.Save(pngTarget, ImageFormat.Png);

                            ImageData imageData = ImageDataFactory.Create(pngTarget.Trim());
                            Image image2 = new Image(imageData);

                            image2.SetWidth(this.outPdf.GetDefaultPageSize().GetWidth() - 50);
                            image2.SetAutoScaleHeight(true);

                            
                            int currentPageBreaks = this.pageBreaks;

                            for (int pageNum = 0; pageNum < pagesNum - currentPageBreaks; pageNum++)
                            {
                                document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                                this.pageBreaks++;
                            }

                            //if (pageBreaks > 0)
                            //{
                            //    this.pageBreaks--;
                            //}

                            document.Add(image2);

                            File.Delete(pngTarget);
                        }
                        catch (System.Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                }
            }
            document.Flush();
            doc.Close(Type.Missing, Type.Missing, Type.Missing);
            app.Quit(Type.Missing, Type.Missing, Type.Missing);

            this.AddPageTitle(this.outPdf,
            pagesNum + 1,
            TitleA + " " + docPath.Trim() + " (1 " + TitleB + ")");
        }

        public static String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyyMMddHHmmssffff");
        }
    }
}
