using System;
using iText.Kernel.Pdf;
using iText.Kernel.Utils;
using System.Runtime.InteropServices;

namespace PDFMergerAndConverter
{
    class Program
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const int SW_SHOW = 5;

        static void Main(string[] args)
        {
            var handle = GetConsoleWindow();

            // Hide Console
            ShowWindow(handle, SW_HIDE);
            ExtractPages(args[0], args[1], Int32.Parse(args[2]), Int32.Parse(args[3]));
        }

        private static void ExtractPages(string sourcePDFpath, string outputPDFpath, int startpage, int endpage)
        {
            var pdfSorce = new PdfDocument(new PdfReader(sourcePDFpath));
            int nEndPage;
            PdfDocument outPdf = new PdfDocument(new PdfWriter(outputPDFpath));
            PdfMerger merger = new PdfMerger(outPdf);

            nEndPage = endpage;
            if (pdfSorce.GetNumberOfPages() < endpage)
            {
                nEndPage = pdfSorce.GetNumberOfPages();
            }

            merger.Merge(pdfSorce, startpage, nEndPage);
            pdfSorce.Close();
            outPdf.Close();

        }
    }
}
