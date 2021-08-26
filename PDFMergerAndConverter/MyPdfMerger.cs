using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Utils;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using System;
using System.Diagnostics;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace PDFMergerAndConverter
{
    class MyPdfMerger
    {
        private PdfDocument outPdf;
        private PdfMerger merger;
        private bool hideShowTitle;
        const string TitleA = "שם המסמך";
        const string TitleB = "עמודים";

        public MyPdfMerger(string outputPDFpath, bool hideShowTitle)
        {
            this.outPdf = new PdfDocument(new PdfWriter(outputPDFpath, new WriterProperties().SetFullCompressionMode(true)));
            this.merger = new PdfMerger(this.outPdf);
            this.hideShowTitle = hideShowTitle;
        }

        public string addPdf(string source, string title)
        {
            if (File.Exists(source.Trim()))
            {
                var pdfSorce = new PdfDocument(new PdfReader(source.Trim()));
                pdfSorce.SetDefaultPageSize(PageSize.LEGAL);
                merger.Merge(pdfSorce, 1, pdfSorce.GetNumberOfPages());

                this.AddPageTitle(this.outPdf,
                                  this.outPdf.GetNumberOfPages() - pdfSorce.GetNumberOfPages() + 1,
                                  title,
                                  pdfSorce.GetNumberOfPages());

                pdfSorce.Close();
            }
            return "";
        }

        private void AddPageTitle(PdfDocument pdfDoc, int pageNumber, string docName, int pages)
        {

            //string pageTitle = TitleA + " " + docName.Trim() + " (" + pages + " " + TitleB + ")";

            //System.Drawing.Bitmap image = ConvertTextToImage(pageTitle, "Arial", 10, System.Drawing.Color.Gray,
            //                                     System.Drawing.Color.White,
            //                                     (int)this.outPdf.GetDefaultPageSize().GetWidth() - 80,
            //                                     20);

            //string tempPath = System.IO.Path.GetTempPath();
            //tempPath += @"/pdf_" + GetTimestamp(new DateTime());

            //Document document = new Document(pdfDoc);
            //var pngTarget = System.IO.Path.ChangeExtension(tempPath, "bmp");
            //image.Save(pngTarget, ImageFormat.Bmp);

            //ImageData imageData = ImageDataFactory.Create(pngTarget.Trim());
            //Image image2 = new Image(imageData);
            //image2.SetWidth(this.outPdf.GetDefaultPageSize().GetWidth() - 50);
            //image2.SetAutoScaleHeight(true);
            //int pagesNum = this.outPdf.GetNumberOfPages();
            //document.Add(image2);
            //File.Delete(pngTarget);


            // string docNameConv = convertStringToUtf8(docName);
            string pageTitle = TitleA + " " + docName.Trim() + " )" + pages + " " + TitleB + "(";
            string ConvertedTitle = ReverseOnlyHebrew(pageTitle);
            if (!this.hideShowTitle)
            {
                return;
            }

            float fontSizeF = 20;
            float allowedWidth = 185;

            try
            {
                string rootFolder = System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
                PdfPage page = pdfDoc.GetPage(pageNumber);
                Rectangle pageSize = page.GetPageSizeWithRotation();
                PdfCanvas pdfCanvas = new PdfCanvas(this.outPdf, pageNumber);
                Rectangle rectangle = new Rectangle(allowedWidth + (pageSize.GetY() * 10), fontSizeF * 3 + (pageSize.GetY() * 10));
                Canvas canvas = new Canvas(pdfCanvas, rectangle);

                PdfFont font = PdfFontFactory.CreateFont(rootFolder + @"\Arial.ttf", PdfEncodings.IDENTITY_H, true);
                Text title =
                  new Text(ConvertedTitle).SetFont(font);

                Paragraph paragraph = new Paragraph().Add(title);
                //paragraph.SetBaseDirection(BaseDirection.RIGHT_TO_LEFT);
                //paragraph.SetFontScript(iText.IO.Util.UnicodeScript.HEBREW);
                paragraph.SetBackgroundColor(ColorConstants.GRAY, 0.7f);
                paragraph.SetFontColor(ColorConstants.WHITE);
                //paragraph.SetWidth(page.GetPageSize().GetWidth() - 30);
                paragraph.SetWidth(page.GetPageSize().GetWidth() + (page.GetPageSize().GetY() * 230));
                paragraph.SetFontSize(fontSizeF - 7 + (page.GetPageSize().GetY() * 4));
                //paragraph.SetBorder(new SolidBorder(1));
                paragraph.SetMargin(3);
                //paragraph.SetMultipliedLeading(1);

                float xCoord = page.GetPageSize().GetWidth() / 2 + (page.GetPageSize().GetY() * 120);
                float yCoord = page.GetPageSize().GetHeight() - 35 - (page.GetPageSize().GetY() * 100);


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

        public void AddImageToPdf(string imageSource, string title, bool firstPage, int pages)
        {
            string tempPath = System.IO.Path.GetTempPath();
            tempPath += @"/pdf_" + GetTimestamp(new DateTime());
            PdfDocument newPage = new PdfDocument((new PdfWriter(tempPath)));
            Document document = new Document(newPage);
            ImageData imageData = ImageDataFactory.Create(imageSource.Trim());
            Image image = new Image(imageData);
            image.SetWidth(this.outPdf.GetDefaultPageSize().GetWidth() - 50);
            image.SetAutoScaleHeight(true);
            int pagesNum = this.outPdf.GetNumberOfPages();
            document.Add(image);
            document.Close();
            this.addPdf(tempPath, title);
            File.Delete(tempPath);
            if (firstPage)
            {
                this.AddPageTitle(this.outPdf, pagesNum + 1, title, pages);
            }

            //    Document document = new Document(this.outPdf);
            //    ImageData imageData = ImageDataFactory.Create(imageSource.Trim());
            //    Image image = new Image(imageData);
            //    image.SetWidth(this.outPdf.GetDefaultPageSize().GetWidth() - 50);
            //    image.SetAutoScaleHeight(true);

            //    int pagesNum = this.outPdf.GetNumberOfPages();
            //    int currentPageBreaks = this.pageBreaks;

            //    this.outPdf.AddNewPage();
            //    for (int pageNum = 0; pageNum < pagesNum - currentPageBreaks; pageNum++)
            //    {
            //        document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
            //        this.pageBreaks++;
            //    }

            //    if (pageBreaks > 0)
            //    {
            //        this.pageBreaks--;
            //    }

            //    document.Add(image);
            //    if (firstPage)
            //    {
            //        this.AddPageTitle(this.outPdf, pagesNum + 1, title, pages);
            //    }

        }
        public void AddTextToPdf(string textSource, string title)
        {
            if (File.Exists(textSource.Trim()))
            {
                string text = File.ReadAllText(textSource.Trim());

                System.Drawing.Bitmap image = ConvertTextToImage(text, "Arial", 10, System.Drawing.Color.White, 
                                                                 System.Drawing.Color.Black, 
                                                                 (int)this.outPdf.GetDefaultPageSize().GetWidth() - 50,
                                                                 (int)this.outPdf.GetDefaultPageSize().GetHeight() - 50);

                string tempPath = System.IO.Path.GetTempPath();
                tempPath += @"/image_" + GetTimestamp(new DateTime());
                var pngTarget = System.IO.Path.ChangeExtension(tempPath, "bmp");
                image.Save(pngTarget, ImageFormat.Bmp);

                this.AddImageToPdf(pngTarget.Trim(), title, true, 1);

                File.Delete(pngTarget);
            }
        }
        public void AddWordToPdf(string docPath, string title)
        {
            var app = new Microsoft.Office.Interop.Word.Application();
            app.Visible = true;
            int pagesNum = this.outPdf.GetNumberOfPages();
            var doc = app.Documents.Open(FileName:docPath.Trim(),Visible:false);
            var firstPage = true;
            //Document document = new Document(this.outPdf, this.outPdf.GetDefaultPageSize(), false);
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


                            this.AddImageToPdf(pngTarget.Trim(), title, firstPage, pane.Pages.Count);
                            firstPage = false;

                            File.Delete(pngTarget);
                        }
                        catch (System.Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                }
            }

            doc.Close(Type.Missing, Type.Missing, Type.Missing);
            app.Quit(Type.Missing, Type.Missing, Type.Missing);
        }

        public static String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyyMMddHHmmssffff");
        }

        public static string convertStringToUtf8(string text)
        {
            Encoding wind1252 = Encoding.GetEncoding(1252);
            Encoding utf8 = Encoding.UTF8;
            byte[] wind1252Bytes = wind1252.GetBytes(text);
            byte[] utf8Bytes = Encoding.Convert(wind1252, utf8, wind1252Bytes);
            string utf8String = Encoding.UTF8.GetString(utf8Bytes);

            return utf8String;
        }

        static public string ReverseOnlyHebrew(string str)
        {   
            string[] arrSplit;
            if (str != null && str != "")
            {
                arrSplit = Regex.Split(str, "( )|([א-ת]+)");
                str = "";
                int arrlenth = arrSplit.Length - 1;
                for (int i = arrlenth; i >= 0; i--)
                {
                    if (arrSplit[i] == " ")
                    {
                        str += " ";
                    }
                    else
                    {
                        if (arrSplit[i] != "")
                        {
                            int outInt;
                            if (int.TryParse(arrSplit[i], out outInt))
                            {
                                str += Convert.ToInt32(arrSplit[i]);
                            }
                            else
                            {
                                arrSplit[i] = arrSplit[i].Trim();
                                byte[] codes = ASCIIEncoding.Default.GetBytes(arrSplit[i].ToCharArray(), 0, 1);
                                if (codes[0] > 47 && codes[0] < 58 || codes[0] > 64 && codes[0] < 91 || codes[0] > 96 && codes[0] < 123)//EDIT 3.1 reverse just hebrew words                              
                                {
                                    str += arrSplit[i].Trim();
                                }
                                else
                                {
                                    str += Reverse(arrSplit[i]);
                                }
                            }
                        }
                    }
                }
            }
            return str;
        }
        static public string Reverse(string str)
        {
            char[] strArray = str.ToCharArray();
            Array.Reverse(strArray);
            return new string(strArray);
        }

        public System.Drawing.Bitmap ConvertTextToImage(string txt, string fontname, int fontsize, System.Drawing.Color bgcolor, System.Drawing.Color fcolor, int width, int Height)
        {
            System.Drawing.Bitmap bmp = new System.Drawing.Bitmap(width, Height);
            using (System.Drawing.Graphics graphics = System.Drawing.Graphics.FromImage(bmp))
            {

                System.Drawing.Font font = new System.Drawing.Font(fontname, fontsize);
                graphics.FillRectangle(new System.Drawing.SolidBrush(bgcolor), 0, 0, bmp.Width, bmp.Height);
                graphics.DrawString(txt, font, new System.Drawing.SolidBrush(fcolor), 0, 0);
                graphics.Flush();
                font.Dispose();
                graphics.Dispose();
            }
            return bmp;
        }
    }
}
