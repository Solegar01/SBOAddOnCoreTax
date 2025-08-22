using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace SBOAddonCoreTax.Services
{
    public static class ExportService
    {
        // ==================== Export XML ====================
        public static bool ExportXml(TaxInvoiceBulk invoice)
        {
            if (invoice == null) throw new ArgumentNullException(nameof(invoice));

            string filePath = GetSaveFilePath("XML Files (*.xml)|*.xml", "TaxInvoice.xml");
            if (string.IsNullOrEmpty(filePath)) return false;

            XmlSerializer serializer = new XmlSerializer(typeof(TaxInvoiceBulk));
            using (var writer = new StreamWriter(filePath))
            {
                serializer.Serialize(writer, invoice);
            }
            return true;
        }

        // ==================== Export CSV ====================
        public static bool ExportCsv(TaxInvoiceBulk invoice)
        {
            if (invoice == null) throw new ArgumentNullException(nameof(invoice));

            string filePath = GetSaveFilePath("CSV Files (*.csv)|*.csv", "TaxInvoice.csv");
            if (string.IsNullOrEmpty(filePath)) return false;

            var sb = new StringBuilder();

            // CSV Header
            sb.AppendLine("TaxInvoiceDate,TaxInvoiceOpt,TrxCode,AddInfo,CustomDoc,CustomDocMonthYear,FacilityStamp,SellerIDTKU,BuyerTin,BuyerDocument,BuyerCountry,BuyerEmail,BuyerIDTKU,GoodServiceOpt,GoodServiceCode,GoodServiceName,Unit,Price,Qty,TotalDiscount,TaxBase,OtherTaxBase,VATRate,VAT,STLGRate,STLG");

            foreach (var taxInvoice in invoice.ListOfTaxInvoice.TaxInvoiceCollection)
            {
                foreach (var gs in taxInvoice.ListOfGoodService.GoodServiceCollection)
                {
                    sb.AppendLine($"{taxInvoice.TaxInvoiceDate},{taxInvoice.TaxInvoiceOpt},{taxInvoice.TrxCode},{taxInvoice.AddInfo},{taxInvoice.CustomDoc},{taxInvoice.CustomDocMonthYear},{taxInvoice.FacilityStamp},{taxInvoice.SellerIDTKU},{taxInvoice.BuyerTin},{taxInvoice.BuyerDocument},{taxInvoice.BuyerCountry},{taxInvoice.BuyerEmail},{taxInvoice.BuyerIDTKU},{gs.Opt},{gs.Code},{gs.Name},{gs.Unit},{gs.Price},{gs.Qty},{gs.TotalDiscount},{gs.TaxBase},{gs.OtherTaxBase},{gs.VATRate},{gs.VAT},{gs.STLGRate},{gs.STLG}");
                }
            }

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
            return true;
        }

        // ==================== Helper: SaveFileDialog in STA thread ====================
        private static string GetSaveFilePath(string filter, string defaultFileName)
        {
            string filePath = null;

            // Append datetime to default file name: TaxInvoice_20250820_083045.xml
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string fileNameWithTimestamp = Path.GetFileNameWithoutExtension(defaultFileName) + "_" + timestamp + Path.GetExtension(defaultFileName);

            Thread t = new Thread(() =>
            {
                using (var saveDialog = new SaveFileDialog
                {
                    Filter = filter,
                    Title = "Save File",
                    FileName = fileNameWithTimestamp
                })
                {
                    if (saveDialog.ShowDialog() == DialogResult.OK)
                        filePath = saveDialog.FileName;
                }
            });

            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();

            return filePath;
        }
    }
}
