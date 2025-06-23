using System.Drawing.Printing;
using System.Drawing.Imaging;
using Microsoft.Reporting.WebForms;
using System.IO;
using System.Collections.Generic;
using System;
using System.Drawing;
using System.Data;

namespace Report
{
    public static class LocalReportExtensions
    {

        
        public static void PrintToPrinter(this LocalReport report, string printerName)
        {
            if (report != null && report.DataSources.Count > 0)
            {
                DataTable dataTable = report.DataSources[0].Value as DataTable; // Assuming first data source is a DataTable

                if (dataTable != null)
                {
                    foreach (DataRow row in dataTable.Rows)
                    {
                        PageSettings pageSettings = new PageSettings();
                        pageSettings.PaperSize = report.GetDefaultPageSettings().PaperSize;
                        pageSettings.Landscape = report.GetDefaultPageSettings().IsLandscape;
                        pageSettings.Margins = report.GetDefaultPageSettings().Margins;
                        pageSettings.PrinterSettings.PrinterName = printerName;


                        Print(report, pageSettings);
                    } 
                }
            }
        }
        private static void Print(LocalReport report, PageSettings pageSettings)
        {
            string deviceInfo =
                $@"<DeviceInfo>
            <OutputFormat>EMF</OutputFormat>
            <PageWidth>{pageSettings.PaperSize.Width / 100.0}in</PageWidth>
            <PageHeight>{pageSettings.PaperSize.Height / 100.0}in</PageHeight>
            <MarginTop>{pageSettings.Margins.Top / 100.0}in</MarginTop>
            <MarginLeft>{pageSettings.Margins.Left / 100.0}in</MarginLeft>
            <MarginRight>{pageSettings.Margins.Right / 100.0}in</MarginRight>
            <MarginBottom>{pageSettings.Margins.Bottom / 100.0}in</MarginBottom>
                </DeviceInfo>";


            
            Warning[] warnings;
            var streams = new List<Stream>();
            var pageIndex = 0;

            report.Render("Image", deviceInfo,
                (name, fileNameExtension, encoding, mimeType, willSeek) =>
                {
                    MemoryStream stream = new MemoryStream();
                    streams.Add(stream);
                    return stream;
                }, out warnings);

            foreach (Stream stream in streams)
                stream.Position = 0;

            if (streams == null || streams.Count == 0)
                throw new Exception("No stream to print.");

            using (PrintDocument printDocument = new PrintDocument())
            {

                PrinterSettings settings = new PrinterSettings
                {
                    PrinterName = pageSettings.PrinterSettings.PrinterName,
                    Copies = 1,
                    Duplex = Duplex.Simplex,
                    PrintRange = PrintRange.AllPages
                };
                printDocument.DefaultPageSettings = new PageSettings
                {
                    PrinterSettings = settings
                    
                };

                printDocument.DefaultPageSettings = pageSettings;
                printDocument.PrinterSettings = settings;

                if (!printDocument.PrinterSettings.IsValid)
                    throw new Exception("Can't find the default printer.");

                printDocument.PrintPage += (sender, e) =>
                {
                    Metafile pageImage = new Metafile(streams[pageIndex]);
                    Rectangle adjustedRect = new Rectangle(
                        e.PageBounds.Left - (int)e.PageSettings.HardMarginX,
                        e.PageBounds.Top - (int)e.PageSettings.HardMarginY,
                        e.PageBounds.Width,
                        e.PageBounds.Height);

                    // Adjusting the rectangle size to fit the actual page size
                    if (pageSettings.Landscape)
                    {
                        adjustedRect = new Rectangle(
                            e.PageBounds.Top - (int)e.PageSettings.HardMarginY,
                            e.PageBounds.Left - (int)e.PageSettings.HardMarginX,
                            e.PageBounds.Height,
                            e.PageBounds.Width);
                    }

                    e.Graphics.FillRectangle(Brushes.White, adjustedRect);
                    e.Graphics.DrawImage(pageImage, adjustedRect);

                    pageIndex++;
                    e.HasMorePages = (pageIndex < streams.Count);
                };

                //printDocument.PrintPage += (sender, e) =>
                //{

                //    Metafile pageImage = new Metafile(streams[pageIndex]);
                //    Rectangle adjustedRect = new Rectangle(e.PageBounds.Left - (int)e.PageSettings.HardMarginX, e.PageBounds.Top - (int)e.PageSettings.HardMarginY, e.PageBounds.Width, e.PageBounds.Height);
                //    e.Graphics.FillRectangle(Brushes.White, adjustedRect);
                //    e.Graphics.DrawImage(pageImage, adjustedRect);
                //    pageIndex++;
                //    e.HasMorePages = (pageIndex < streams.Count);
                //    e.Graphics.DrawRectangle(Pens.Red, adjustedRect);
                //};



                printDocument.EndPrint += (sender, e) =>
                {
                    foreach (Stream stream in streams)
                        stream.Close();
                    streams = null;
                };

                printDocument.Print();
            }
        }
        
    }
}
