using System;
using iText.Kernel.Pdf;
using iText.Kernel.Utils;
using System.Runtime.InteropServices;
using System.IO;

namespace PDFMergerAndConverter
{
    class Program
    {
        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow(); 

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        const int SW_HIDE = 0;
        const string PDF_T = "[PDF]";
        const string PNG_T = "[PNG]";
        const string JPG_T = "[JPG]";
        const string TIF_T = "[TIF]";
        const string BMP_T = "[BMP]";
        const string TXT_T = "[TXT]";
        const string DOC_T = "[DOC]";
        const int SW_SHOW = 5;
        static MyPdfMerger MypdfMerger;
        static StreamReader PdfPathsReader;
        static bool HideShowTitle = true;

        static int Main(string[] args)
        {   
            // Agr 1 = text file with file paths
            // Agr 2 = Path for new PDF file

            var handle = GetConsoleWindow();
            string fileLine;

            // Hide Console
            ShowWindow(handle, SW_HIDE);
            if(args.Length < 2 || args[0] == string.Empty || args[1] == string.Empty)
            {
                return -1;
            }
  
            if (!File.Exists(args[0].Trim()))
            {
                return -2;
            }

            if ( args.Length >= 3)
            {   
                HideShowTitle = bool.Parse(args[2].Split("=")[1]);
            }

            MypdfMerger = new MyPdfMerger(args[1], HideShowTitle);
            PdfPathsReader = new StreamReader(args[0]);

            while ((fileLine = PdfPathsReader.ReadLine()) != null)
            {
                if (fileLine.Length >= 5)
                {
                    switch (fileLine.Substring(0, 5))
                    {
                        case (PDF_T):
                            MypdfMerger.addPdf(fileLine.Substring(5).Split(",")[0], fileLine.Substring(5).Split(",")[1]);
                            break;
                        case (PNG_T): // Acting like OR
                        case (JPG_T):
                        case (TIF_T):
                        case (BMP_T):
                            MypdfMerger.AddImageToPdf(fileLine.Substring(5).Split(",")[0], fileLine.Substring(5).Split(",")[1], true);
                            break;
                        case (TXT_T):
                            MypdfMerger.AddTextToPdf(fileLine.Substring(5).Split(",")[0], fileLine.Substring(5).Split(",")[1]);
                            break;
                        case (DOC_T):
                            MypdfMerger.AddWordToPdf(fileLine.Substring(5).Split(",")[0], fileLine.Substring(5).Split(",")[1]);
                            break;
                        //default:

                    }
                }
            }

            MypdfMerger.ClosePdf();

            return 0;
        }

    }
}
