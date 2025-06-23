using Microsoft.Reporting.WebForms;
using Newtonsoft.Json;
using OfficeOpenXml;
using Report.Report.Model;
using Report.Service;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using ZXing;
using PdfiumViewer;

namespace Report.Repository
{
    public class PostCode
    {
        public string Postcode { get; set; }
        public string Route { get; set; }
    }
    public class PrintLabel_REPO : IPrintLabel
    {
        readonly private LogWriter _logWriter = new LogWriter();
        public string Print_label(string pathjson, string pathRDLC, string printername)
        {

            DataTable dtprintList = ConvertJson(pathjson);
            PageSettings pageSettings = new PageSettings();

            foreach (DataRow row in dtprintList.Rows)
            {
                DataTable singleRowTable = dtprintList.Clone();
                singleRowTable.ImportRow(row);


                ReportViewer InvoiceReport = new ReportViewer();
                InvoiceReport.LocalReport.ReportPath = pathRDLC;//@"C:\work\Report\Report\bin\Debug\Report\rptPrintLabel.rdlc";

                ReportDataSource DataSet1 = new ReportDataSource()
                {
                    Name = "PrintLabel",
                    Value = singleRowTable,
                };

                InvoiceReport.LocalReport.DataSources.Clear();
                InvoiceReport.LocalReport.DataSources.Add(DataSet1);

                LocalReport lr = InvoiceReport.LocalReport;
                lr.PrintToPrinter(printername); //"Honeywell PC42t plus (203 dpi) (Copy 1)"

                _logWriter.LogWrite(string.Format("print"));
                //PrintReport(lr, pageSettings);

                lr.Dispose();
            }
            return "ok";
            //throw new NotImplementedException();
        }

        private DataTable ConvertJson(string pathjson)
        {
            string filePath = pathjson;//@"C:\Users\TDumrongmun\Desktop\k-pro-api-mobile\KWMSAPI\APP_DATA\files\SSANG\Label\S_SANG_PrintLabel.json";

            //string PostCodePath = Path.Combine(Directory.GetCurrentDirectory(), @"Report\File\PostCodeTP.xlsx");
            string PostCodePath = Path.Combine(Directory.GetCurrentDirectory(), "AppData", "File", "PostCodeTP.xlsx");
            //string PostCodePath = Path.Combine(appDataPath, "\\PostCodeTP.xlsx\\");
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            var postCodes = new List<PostCode>();



            using (var package = new ExcelPackage(new FileInfo(PostCodePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet != null)
                {
                    for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                    {
                        PostCode post = new PostCode();

                        // Read data from the Excel worksheet
                        post.Postcode = worksheet.Cells[row, 1].Value?.ToString() ?? string.Empty;
                        post.Route = worksheet.Cells[row, 2].Value?.ToString() ?? string.Empty;

                        postCodes.Add(post);
                    }
                }
            }
            try
            {
                if (File.Exists(filePath))
                {
                    string jsonContent = File.ReadAllText(filePath);

                    dsPrintLabel.dtPrintLabelDataTable printList = JsonConvert.DeserializeObject<dsPrintLabel.dtPrintLabelDataTable>(jsonContent);


                    var barcodeWriter = new BarcodeWriter();
                    barcodeWriter.Format = BarcodeFormat.CODE_128;
                    barcodeWriter.Options.Width = 200;
                    barcodeWriter.Options.Height = 200;
                    barcodeWriter.Options.Margin = 0;
                    barcodeWriter.Options.PureBarcode = true;



                    //dsPrintLabel.dtPrintLabelDataTable pdata = new dsPrintLabel.dtPrintLabelDataTable();
                    foreach (var i in printList)
                    {
                        dsPrintLabel.dtPrintLabelDataTable c = new dsPrintLabel.dtPrintLabelDataTable();

                        if (i.user_defined_16 == "02")
                        {
                            var barcodeBitmapPL = barcodeWriter.Write(i.tacking_code);
                            byte[] barcodeBytesPL = BitmapToByteArray(barcodeBitmapPL);
                            i.barcode = Convert.ToBase64String(barcodeBytesPL);
                        }
                        else if (i.user_defined_16 == "04")
                        {
                            var barcodeBitmapPL = barcodeWriter.Write(i.tacking_code);
                            byte[] barcodeBytesPL = BitmapToByteArray(barcodeBitmapPL);
                            i.barcode = Convert.ToBase64String(barcodeBytesPL);
                            i.hub_code = postCodes.Where(x => x.Postcode == i.St_Zip).Select(x => x.Route).FirstOrDefault().ToString();
                        }

                        else if (i.user_defined_16 == "FZ-REC")
                        {
                            var barcodeWriterQR = new BarcodeWriter();
                            barcodeWriter.Format = BarcodeFormat.QR_CODE;
                            barcodeWriter.Options.Width = 200;
                            barcodeWriter.Options.Height = 200;
                            barcodeWriter.Options.Margin = 0;
                            barcodeWriter.Options.PureBarcode = true;

                            var barcodeBitmapFZ = barcodeWriter.Write(i.sku);
                            byte[] barcodeBytesPL = BitmapToByteArray(barcodeBitmapFZ);
                            i.barcode = Convert.ToBase64String(barcodeBytesPL);

                        }
                        else
                        {
                            var barcodeBitmap = barcodeWriter.Write(i.inv_ref);
                            byte[] barcodeBytes = BitmapToByteArray(barcodeBitmap);
                            i.barcode = Convert.ToBase64String(barcodeBytes);
                        }
                        //pdata.(c);
                    }

                    return printList;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                throw;
            }
            throw new NotImplementedException();
        }

        public string Print_label_PDF(string pathJson, string pathRDLC, string printerName)
        {
            DataTable dtprintList = ConvertJsonPDF(pathJson);
            foreach (DataRow row in dtprintList.Rows)
            {
                string filePath = row["Path_Pdf"].ToString();
                if (File.Exists(filePath))
                {
                    using (var pdfDocument = PdfDocument.Load(filePath))
                    {
                        var printDocument = pdfDocument.CreatePrintDocument();
                        printDocument.PrinterSettings.PrinterName = printerName;
                        printDocument.PrinterSettings.DefaultPageSettings.PrinterResolution.Kind = PrinterResolutionKind.High;
                        printDocument.Print();
                    }
                }
            }
            return "ok";
        }

        private DataTable ConvertJsonPDF(string pathjson)
        {
            string filePath = pathjson;//@"C:\Users\TDumrongmun\Desktop\k-pro-api-mobile\KWMSAPI\APP_DATA\files\SSANG\Label\S_SANG_PrintLabel.json";

            try
            {
                if (File.Exists(filePath))
                {
                    string jsonContent = File.ReadAllText(filePath);
                    dsPrintLabel.dtPrintLabelDataTable printList = JsonConvert.DeserializeObject<dsPrintLabel.dtPrintLabelDataTable>(jsonContent);

                    return printList;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                throw;
            }
            throw new NotImplementedException();
        }

        private static byte[] BitmapToByteArray(Bitmap bitmap)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {

                bitmap.Save(memoryStream, ImageFormat.Png);

                return memoryStream.ToArray();
            }
        }
    }
    public class LogWriter //: ILogWriter
    {
        private string m_exePath = string.Empty;
        //public LogWriter(string logMessage)
        //{
        //    LogWrite(logMessage);
        //}
        public void LogWrite(string logMessage)
        {
            m_exePath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            var filePathBatch = Path.Combine(Directory.GetCurrentDirectory(), @"AppData\");
            try
            {
                using (StreamWriter w = File.AppendText(filePathBatch + "\\Log\\" + $"log_{DateTime.Now.ToString("yyyyMMdd")}.txt"))
                {
                    Log(logMessage, w);
                }
            }
            catch (Exception ex)
            {
            }
        }

        private void Log(string logMessage, TextWriter txtWriter)
        {
            try
            {
                txtWriter.Write("\r\nLog Entry : ");
                txtWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(),
                    DateTime.Now.ToLongDateString());
                txtWriter.WriteLine("  :");
                txtWriter.WriteLine("  :{0}", logMessage);
                txtWriter.WriteLine("-------------------------------");
            }
            catch (Exception ex)
            {
            }
        }

    }

}
