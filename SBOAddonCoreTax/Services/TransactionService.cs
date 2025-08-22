using SAPbobsCOM;
using SBOAddonCoreTax.Models;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SBOAddonCoreTax.Services
{
    public static class TransactionService
    {
        public static bool AddDataCoretax(CoretaxModel model)
        {
            try
            {
                Company oCompany = CompanyService.GetCompany();
                SAPbobsCOM.CompanyService oCompanyService = oCompany.GetCompanyService();
                int seriesId = GetSeriesIdCoretax(oCompany);
                // 1. Get GeneralService for UDO
                GeneralService oGeneralService = oCompanyService.GetGeneralService("T2_CORETAX");

                // 2. Create header
                GeneralData oGeneralData = (GeneralData)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);
                
                // --- Header fields ---
                oGeneralData.SetProperty("Series", seriesId);
                oGeneralData.SetProperty("U_T2_Doc_Date", DateTime.Now);
                oGeneralData.SetProperty("U_T2_AR_INV", model.IsARInvoice ? "Y" : "N");
                oGeneralData.SetProperty("U_T2_AR_DP", model.IsARDownPayment ? "Y" : "N");
                oGeneralData.SetProperty("U_T2_AR_CM", model.IsARCreditMemo ? "Y" : "N");

                if (model.FromDate != null) oGeneralData.SetProperty("U_T2_From_Date", model.FromDate);
                if (model.ToDate != null) oGeneralData.SetProperty("U_T2_To_Date", model.ToDate);
                if (model.FromDocNum > 0) oGeneralData.SetProperty("U_T2_From_Doc", model.FromDocNum);
                if (model.ToDocNum > 0) oGeneralData.SetProperty("U_T2_To_Doc", model.ToDocNum);
                if (model.FromDocEntry > 0) oGeneralData.SetProperty("U_T2_From_Doc_Entry", model.FromDocEntry);
                if (model.ToDocEntry > 0) oGeneralData.SetProperty("U_T2_To_Doc_Entry", model.ToDocEntry);

                if (!string.IsNullOrEmpty(model.FromCust)) oGeneralData.SetProperty("U_T2_From_Cust", model.FromCust);
                if (!string.IsNullOrEmpty(model.ToCust)) oGeneralData.SetProperty("U_T2_To_Cust", model.ToCust);
                if (!string.IsNullOrEmpty(model.FromBranch)) oGeneralData.SetProperty("U_T2_From_Branch", model.FromBranch);
                if (!string.IsNullOrEmpty(model.ToBranch)) oGeneralData.SetProperty("U_T2_To_Branch", model.ToBranch);
                if (!string.IsNullOrEmpty(model.FromOutlet)) oGeneralData.SetProperty("U_T2_From_Outlet", model.FromOutlet);
                if (!string.IsNullOrEmpty(model.ToOutlet)) oGeneralData.SetProperty("U_T2_To_Outlet", model.ToOutlet);

                // 3. Add detail lines
                GeneralDataCollection oChildren = oGeneralData.Child("T2_CORETAX_DT");
                foreach (var line in model.Detail)
                {
                    GeneralData oChild = oChildren.Add();

                    if (!string.IsNullOrEmpty(line.TIN)) oChild.SetProperty("U_T2_TIN", line.TIN);
                    if (!string.IsNullOrEmpty(line.DocEntry)) oChild.SetProperty("U_T2_DocEntry", line.DocEntry);
                    if (!string.IsNullOrEmpty(line.LineNum)) oChild.SetProperty("U_T2_LineNum", line.LineNum);
                    if (!string.IsNullOrEmpty(line.InvDate)) oChild.SetProperty("U_T2_Inv_Date", line.InvDate);
                    if (!string.IsNullOrEmpty(line.NoDocument)) oChild.SetProperty("U_T2_No_Doc", line.NoDocument);
                    if (!string.IsNullOrEmpty(line.ObjectType)) oChild.SetProperty("U_T2_Object_Type", line.ObjectType);
                    if (!string.IsNullOrEmpty(line.ObjectName)) oChild.SetProperty("U_T2_Object_Name", line.ObjectName);
                    if (!string.IsNullOrEmpty(line.BPCode)) oChild.SetProperty("U_T2_BP_Code", line.BPCode);
                    if (!string.IsNullOrEmpty(line.BPName)) oChild.SetProperty("U_T2_BP_Name", line.BPName);
                    if (!string.IsNullOrEmpty(line.SellerIDTKU)) oChild.SetProperty("U_T2_Seller_IDTKU", line.SellerIDTKU);
                    if (!string.IsNullOrEmpty(line.BuyerDocument)) oChild.SetProperty("U_T2_Buyer_Doc", line.BuyerDocument);
                    if (!string.IsNullOrEmpty(line.NomorNPWP)) oChild.SetProperty("U_T2_Nomor_NPWP", line.NomorNPWP);
                    if (!string.IsNullOrEmpty(line.NPWPName)) oChild.SetProperty("U_T2_NPWP_Name", line.NPWPName);
                    if (!string.IsNullOrEmpty(line.NPWPAddress)) oChild.SetProperty("U_T2_NPWP_Address", line.NPWPAddress);
                    if (!string.IsNullOrEmpty(line.BuyerIDTKU)) oChild.SetProperty("U_T2_Buyer_IDTKU", line.BuyerIDTKU);
                    if (!string.IsNullOrEmpty(line.ItemCode)) oChild.SetProperty("U_T2_Item_Code", line.ItemCode);
                    if (!string.IsNullOrEmpty(line.ItemName)) oChild.SetProperty("U_T2_Item_Name", line.ItemName);
                    if (!string.IsNullOrEmpty(line.ItemUnit)) oChild.SetProperty("U_T2_Item_Unit", line.ItemUnit);

                    // --- numeric fields (pastikan UDF tipe number/price/rate) ---
                    if (line.ItemPrice > 0) oChild.SetProperty("U_T2_Item_Price", line.ItemPrice);
                    if (line.Qty > 0) oChild.SetProperty("U_T2_Qty", line.Qty);
                    if (line.TotalDisc > 0) oChild.SetProperty("U_T2_Total_Disc", line.TotalDisc);
                    if (line.TaxBase > 0) oChild.SetProperty("U_T2_Tax_Base", line.TaxBase);
                    if (line.OtherTaxBase > 0) oChild.SetProperty("U_T2_Other_Tax_Base", line.OtherTaxBase);
                    if (line.VATRate > 0) oChild.SetProperty("U_T2_VAT_Rate", line.VATRate);
                    if (line.AmountVAT > 0) oChild.SetProperty("U_T2_Amount_VAT", line.AmountVAT);
                    if (line.STLGRate > 0) oChild.SetProperty("U_T2_STLG_Rate", line.STLGRate);
                    if (line.STLG > 0) oChild.SetProperty("U_T2_STLG", line.STLG);

                    // --- string fields ---
                    if (!string.IsNullOrEmpty(line.JenisPajak)) oChild.SetProperty("U_T2_Jenis_Pajak", line.JenisPajak);
                    if (!string.IsNullOrEmpty(line.KetTambahan)) oChild.SetProperty("U_T2_Ket_Tambahan", line.KetTambahan);
                    if (!string.IsNullOrEmpty(line.PajakPengganti)) oChild.SetProperty("U_T2_Pajak_Pengganti", line.PajakPengganti);
                    if (!string.IsNullOrEmpty(line.Referensi)) oChild.SetProperty("U_T2_Referensi", line.Referensi);
                    if (!string.IsNullOrEmpty(line.Status)) oChild.SetProperty("U_T2_Status", line.Status);
                    if (!string.IsNullOrEmpty(line.KodeDokumenPendukung)) oChild.SetProperty("U_T2_Kode_Dok_Pendukung", line.KodeDokumenPendukung);
                    if (!string.IsNullOrEmpty(line.Branch)) oChild.SetProperty("U_T2_Branch", line.Branch);
                    if (!string.IsNullOrEmpty(line.AddInfo)) oChild.SetProperty("U_T2_Add_Info", line.AddInfo);
                    if (!string.IsNullOrEmpty(line.BuyerCountry)) oChild.SetProperty("U_T2_Buyer_Country", line.BuyerCountry);
                    if (!string.IsNullOrEmpty(line.BuyerEmail)) oChild.SetProperty("U_T2_Buyer_Email", line.BuyerEmail);
                }

                // 4. Save data
                GeneralDataParams oGeneralParams = oGeneralService.Add(oGeneralData);
                int newDocEntry = Convert.ToInt32(oGeneralParams.GetProperty("DocEntry"));

                return true;
            }
            catch (Exception ex)
            {
                // kalau mau return false aja biar lebih aman
                throw new Exception($"Error AddDataCoretax: {ex.Message}", ex);
            }
        }

        public static CoretaxModel GetCoretaxByKey(int docEntry)
        {
            Company oCompany = CompanyService.GetCompany();

            try
            {
                SAPbobsCOM.CompanyService oCompanyService = oCompany.GetCompanyService();
                
                SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("T2_CORETAX");
                
                SAPbobsCOM.GeneralDataParams oParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(
                    SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oParams.SetProperty("DocEntry", docEntry);
                
                SAPbobsCOM.GeneralData oHeader = oGeneralService.GetByParams(oParams);

                var series = GetSeriesByDocEntry("T2_CORETAX", docEntry);

                var model = new CoretaxModel
                {
                    DocNum = Convert.ToInt32(oHeader.GetProperty("DocNum")),
                    DocDate = oHeader.GetProperty("U_T2_Doc_Date") == null ? null : (DateTime?)Convert.ToDateTime(oHeader.GetProperty("U_T2_Doc_Date")),
                    PostDate = oHeader.GetProperty("U_T2_Posting_Date") == null ? null : (DateTime?)Convert.ToDateTime(oHeader.GetProperty("U_T2_Posting_Date")),
                    SeriesId = series.SeriesId,
                    SeriesName = series.SeriesName,
                    IsARInvoice = (oHeader.GetProperty("U_T2_AR_INV")?.ToString() == "Y"),
                    IsARDownPayment = (oHeader.GetProperty("U_T2_AR_DP")?.ToString() == "Y"),
                    IsARCreditMemo = (oHeader.GetProperty("U_T2_AR_CM")?.ToString() == "Y"),

                    FromDocNum = oHeader.GetProperty("U_T2_From_Doc") == null ? null : (int?)Convert.ToInt32(oHeader.GetProperty("U_T2_From_Doc")),
                    ToDocNum = oHeader.GetProperty("U_T2_To_Doc") == null ? null : (int?)Convert.ToInt32(oHeader.GetProperty("U_T2_To_Doc")),
                    FromDocEntry = oHeader.GetProperty("U_T2_From_Doc_Entry") == null ? null : (int?)Convert.ToInt32(oHeader.GetProperty("U_T2_From_Doc_Entry")),
                    ToDocEntry = oHeader.GetProperty("U_T2_To_Doc_Entry") == null ? null : (int?)Convert.ToInt32(oHeader.GetProperty("U_T2_To_Doc_Entry")),

                    FromDate = oHeader.GetProperty("U_T2_From_Date") == null ? null : (DateTime?)Convert.ToDateTime(oHeader.GetProperty("U_T2_From_Date")),
                    ToDate = oHeader.GetProperty("U_T2_To_Date") == null ? null : (DateTime?)Convert.ToDateTime(oHeader.GetProperty("U_T2_To_Date")),

                    FromCust = oHeader.GetProperty("U_T2_From_Cust")?.ToString(),
                    ToCust = oHeader.GetProperty("U_T2_To_Cust")?.ToString(),
                    FromBranch = oHeader.GetProperty("U_T2_From_Branch")?.ToString(),
                    ToBranch = oHeader.GetProperty("U_T2_To_Branch")?.ToString(),
                    FromOutlet = oHeader.GetProperty("U_T2_From_Outlet")?.ToString(),
                    ToOutlet = oHeader.GetProperty("U_T2_To_Outlet")?.ToString(),

                    Status = oHeader.GetProperty("Status")?.ToString() ?? "O"
                };

                // 6. Map child table (detail lines)
                SAPbobsCOM.GeneralDataCollection children = oHeader.Child("T2_CORETAX1"); // UDO child table name
                foreach (SAPbobsCOM.GeneralData line in children)
                {
                    var detail = new InvoiceDataModel
                    {
                        TIN = line.GetProperty("U_T2_TIN")?.ToString(),
                        DocEntry = line.GetProperty("U_T2_DocEntry")?.ToString(),
                        LineNum = line.GetProperty("U_T2_LineNum")?.ToString(),
                        InvDate = line.GetProperty("U_T2_Inv_Date")?.ToString(),
                        NoDocument = line.GetProperty("U_T2_No_Doc")?.ToString(),
                        ObjectType = line.GetProperty("U_T2_Object_Type")?.ToString(),
                        ObjectName = line.GetProperty("U_T2_Object_Name")?.ToString(),
                        BPCode = line.GetProperty("U_T2_BP_Code")?.ToString(),
                        BPName = line.GetProperty("U_T2_BP_Name")?.ToString(),
                        SellerIDTKU = line.GetProperty("U_T2_Seller_IDTKU")?.ToString(),
                        BuyerDocument = line.GetProperty("U_T2_Buyer_Doc")?.ToString(),
                        NomorNPWP = line.GetProperty("U_T2_Nomor_NPWP")?.ToString(),
                        NPWPName = line.GetProperty("U_T2_NPWP_Name")?.ToString(),
                        NPWPAddress = line.GetProperty("U_T2_NPWP_Address")?.ToString(),
                        BuyerIDTKU = line.GetProperty("U_T2_Buyer_IDTKU")?.ToString(),
                        ItemCode = line.GetProperty("U_T2_Item_Code")?.ToString(),
                        ItemName = line.GetProperty("U_T2_Item_Name")?.ToString(),
                        ItemUnit = line.GetProperty("U_T2_Item_Unit")?.ToString(),
                        ItemPrice = line.GetProperty("U_T2_Item_Price") == null ? 0 : Convert.ToDouble(line.GetProperty("U_ItemPrice")),
                        Qty = line.GetProperty("U_T2_Qty") == null ? 0 : Convert.ToDouble(line.GetProperty("U_Qty")),
                        TotalDisc = line.GetProperty("U_T2_Total_Disc") == null ? 0 : Convert.ToDouble(line.GetProperty("U_TotalDisc")),
                        TaxBase = line.GetProperty("U_T2_Tax_Base") == null ? 0 : Convert.ToDouble(line.GetProperty("U_TaxBase")),
                        OtherTaxBase = line.GetProperty("U_T2_Other_Tax_Base") == null ? 0 : Convert.ToDouble(line.GetProperty("U_OtherTaxBase")),
                        VATRate = line.GetProperty("U_T2_VAT_Rate") == null ? 0 : Convert.ToDouble(line.GetProperty("U_VATRate")),
                        AmountVAT = line.GetProperty("U_T2_Amount_VAT") == null ? 0 : Convert.ToDouble(line.GetProperty("U_AmountVAT")),
                        STLGRate = line.GetProperty("U_T2_STLG_Rate") == null ? 0 : Convert.ToDouble(line.GetProperty("U_STLGRate")),
                        STLG = line.GetProperty("U_T2_STLG") == null ? 0 : Convert.ToDouble(line.GetProperty("U_STLG")),
                        JenisPajak = line.GetProperty("U_T2_Jenis_Pajak")?.ToString(),
                        KetTambahan = line.GetProperty("U_T2_Ket_Tambahan")?.ToString(),
                        PajakPengganti = line.GetProperty("U_T2_Pajak_Pengganti")?.ToString(),
                        Referensi = line.GetProperty("U_T2_Referensi")?.ToString(),
                        Status = line.GetProperty("U_T2_Status")?.ToString(),
                        KodeDokumenPendukung = line.GetProperty("U_T2_KodeDokPendukung")?.ToString(),
                        Branch = line.GetProperty("U_T2_Branch")?.ToString(),
                        AddInfo = line.GetProperty("U_T2_Add_Info")?.ToString(),
                        BuyerCountry = line.GetProperty("U_T2_Buyer_Country")?.ToString(),
                        BuyerEmail = line.GetProperty("U_T2_Buyer_Email")?.ToString()
                    };

                    model.Detail.Add(detail);
                }

                return model;
            }
            catch (Exception ex)
            {
                throw new Exception($"Error while getting Coretax UDO DocEntry={docEntry}: {ex.Message}", ex);
            }
        }

        public static (int SeriesId, string SeriesName) GetSeriesByDocEntry(string udoTable, int docEntry)
        {
            Company oCompany = CompanyService.GetCompany();
            Recordset oRs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                string sql = $@"
                SELECT T0.Series, ISNULL(T1.SeriesName, '') AS SeriesName
                FROM [@{udoTable}] T0
                LEFT JOIN NNM1 T1 ON T0.Series = T1.Series AND T1.ObjectCode = '{udoTable}'
                WHERE T0.DocEntry = {docEntry}";

                oRs.DoQuery(sql);

                if (!oRs.EoF)
                {
                    int seriesId = Convert.ToInt32(oRs.Fields.Item("Series").Value);
                    string seriesName = oRs.Fields.Item("SeriesName").Value.ToString();
                    return (seriesId, seriesName);
                }

                return (0, null);
            }
            catch (Exception ex)
            {
                throw new Exception("Error get series by DocEntry: " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
            }
        }


        private static int GetSeriesIdCoretax(Company oCompany)
        {
            try
            {
                Recordset oRecordset = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                // panggil SQL function
                string sql = "SELECT dbo.T2_GET_SERIES_ID_CORETAX() AS SeriesId";
                oRecordset.DoQuery(sql);

                if (!oRecordset.EoF)
                {
                    return Convert.ToInt32(oRecordset.Fields.Item("SeriesId").Value);
                }
                else
                {
                    throw new Exception("SeriesId not found.");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error GetSeriesIdCoretax2025: {ex.Message}", ex);
            }
        }
    }
}
