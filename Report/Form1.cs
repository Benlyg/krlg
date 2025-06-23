using Microsoft.Reporting.WebForms;
using Newtonsoft.Json;
using Report.Report.Model;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZXing;

namespace Report
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public static DataTable ConvertJson()
        {
            string filePath = @"C:\Users\TDumrongmun\Desktop\k-pro-api-mobile\KWMSAPI\APP_DATA\files\SSANG\Label\S_SANG_PrintLabel.json";

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



                    //dsPrintLabel.dtPrintLabelDataTable pdata = new dsPrintLabel.dtPrintLabelDataTable();
                    foreach (var i in printList)
                    {
                        dsPrintLabel.dtPrintLabelDataTable c = new dsPrintLabel.dtPrintLabelDataTable();
                        var barcodeBitmap = barcodeWriter.Write(i.tacking_code);
                        byte[] barcodeBytes = BitmapToByteArray(barcodeBitmap);
                        i.barcode = Convert.ToBase64String(barcodeBytes);
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
        }
        private static byte[] BitmapToByteArray(Bitmap bitmap)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {

                bitmap.Save(memoryStream, ImageFormat.Png);

                return memoryStream.ToArray();
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            DataTable dtprintList = ConvertJson();
            PageSettings pageSettings = new PageSettings(); 

            foreach (DataRow row in dtprintList.Rows)
            {
                DataTable singleRowTable = dtprintList.Clone(); 
                singleRowTable.ImportRow(row); 

                ReportViewer InvoiceReport = new ReportViewer();
                InvoiceReport.LocalReport.ReportPath = @"C:\work\Report\Report\bin\Debug\Report\rptPrintLabel.rdlc";

                ReportDataSource DataSet1 = new ReportDataSource()
                {
                    Name = "PrintLabel",
                    Value = singleRowTable,
                };

                InvoiceReport.LocalReport.DataSources.Clear(); 
                InvoiceReport.LocalReport.DataSources.Add(DataSet1);

                LocalReport lr = InvoiceReport.LocalReport;
                lr.PrintToPrinter("Honeywell PC42t plus (203 dpi) (Copy 1)");
                //PrintReport(lr, pageSettings);

                lr.Dispose();
            }
        }
        

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dgv.DataSource = bs;
        }

        private void bs_CurrentChanged(object sender, EventArgs e)
        {

        }

        private void dgv_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
