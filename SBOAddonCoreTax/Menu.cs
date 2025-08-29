using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Serialization;
using SAPbobsCOM;
using SAPbouiCOM.Framework;
using SBOAddonCoreTax.Models;
using SBOAddonCoreTax.Services;

namespace SBOAddonCoreTax
{
    class Menu
    {
        CoretaxModel coretaxModel = new CoretaxModel();

        private Dictionary<string, bool> SelectedCkBox = new Dictionary<string, bool>()
            {
                { "OINV", false},
                { "ODPI", false},
                { "ORIN", false},
            };
        
        private bool FilterIsShow = false;
        private List<string> SelectedCardCode = new List<string>();
        private List<FilterDataModel> FindListModel = new List<FilterDataModel>();
        private TaxInvoiceBulk taxInvoiceBulk = new TaxInvoiceBulk();
        private string FORMUID = "";
        SAPbouiCOM.ProgressBar _pb = null;
        private Dictionary<string, string> cbBranchValues = new Dictionary<string, string>();
        private Dictionary<string, string> cbOutletValues = new Dictionary<string, string>();

        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles'

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "SBOAddonCoreTax";
            oCreationPackage.String = "Generate Coretax";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;

            oMenus = oMenuItem.SubMenus;

            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch
            {

            }

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("SBOAddonCoreTax");
                Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
                Application.SBO_Application.RightClickEvent += SBO_Application_RightClickEvent;
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "SBOAddonCoreTax.CoretaxForm";
                oCreationPackage.String = "Generate Coretax";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "SBOAddonCoreTax.CoretaxForm")
                {
                    CoretaxForm activeForm = new CoretaxForm();
                    activeForm.Show();
                }
                if (!pVal.BeforeAction && pVal.MenuUID == "1282")
                {
                    try
                    {
                        // Get the active form
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;

                        // Check if it is YOUR form
                        if (oForm.TypeEx == "SBOAddonCoreTax.CoretaxForm") // <-- replace with your FormTypeEx
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                Task.Run(async () => {
                                    await GetNewModel(oForm);
                                }).ContinueWith(task => MapModelToUI(oForm));
                            }
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

        private void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo EventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true; // Set to true to allow the default right-click menu to appear after your code executes

            // Check the form and item where the right-click occurred
            if (EventInfo.FormUID == FORMUID && EventInfo.BeforeAction)
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(EventInfo.FormUID);
                Task.Run(async () =>
                {
                    try
                    {
                        coretaxModel = await TransactionService.GetCoretaxByKey(10);
                    }
                    catch (Exception)
                    {
                        throw;
                    }
                }).ContinueWith(task =>
                {
                    MapModelToUI(oForm);
                });
                BubbleEvent = false;
                //Application.SBO_Application.SetStatusBarMessage("Right click event",SAPbouiCOM.BoMessageTime.bmt_Short,false);
                // Perform actions based on the right-click context
                // For example, enable/disable specific menu items, add custom menu options, etc.
                // You can also check EventInfo.Row and EventInfo.Col for matrix-specific right-clicks
            }

            // If you want to prevent the default right-click menu from appearing, set BubbleEvent to false
            // BubbleEvent = false;
        }

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (pVal.FormTypeEx == "SBOAddonCoreTax.CoretaxForm")
            {
                try
                {
                    FORMUID = FormUID;
                    
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE &&
                    pVal.BeforeAction == false) // setelah validasi selesai
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                        if(pVal.ItemUID == "TFromDt")
                        {
                            SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("TFromDt").Specific;
                            var oldVal = coretaxModel.FromDate;
                            if (string.IsNullOrWhiteSpace(oEdit.Value))
                            {
                                coretaxModel.FromDate = null;
                            }
                            else if (DateTime.TryParseExact(oEdit.Value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                            {
                                if (parsedDate != coretaxModel.FromDate)
                                {
                                    coretaxModel.FromDate = parsedDate;
                                }
                            }
                            else
                            {
                                coretaxModel.FromDate = null;
                            }
                            if (oldVal != coretaxModel.FromDate)
                            {
                                ResetDetail(oForm);
                            }
                        }
                        if (pVal.ItemUID == "TToDt")
                        {
                            SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("TToDt").Specific;
                            var oldVal = coretaxModel.ToDate;
                            if (string.IsNullOrWhiteSpace(oEdit.Value))
                            {
                                coretaxModel.ToDate = null;
                            }
                            else if (DateTime.TryParseExact(oEdit.Value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                            {
                                if (parsedDate != coretaxModel.ToDate)
                                {
                                    coretaxModel.ToDate = parsedDate;
                                }
                            }
                            else
                            {
                                coretaxModel.ToDate = null;
                            }
                            if (oldVal != coretaxModel.ToDate)
                            {
                                ResetDetail(oForm);
                            }
                        }
                        if (pVal.ItemUID == "TFromDoc")
                        {
                            SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("TFromDoc").Specific;
                            var oldVal = coretaxModel.FromDocNum;

                            if (string.IsNullOrWhiteSpace(oEdit.Value))
                            {
                                // clear model
                                coretaxModel.FromDocNum = null;
                                coretaxModel.FromDocEntry = null;

                                if (oldVal != null) // only reset if value actually changed
                                {
                                    ResetDetail(oForm);
                                }
                            }
                            else if (int.TryParse(oEdit.Value, out int newVal))
                            {
                                if (oldVal == null || newVal != oldVal)
                                {
                                    coretaxModel.FromDocNum = newVal;
                                    ResetDetail(oForm);
                                }
                            }
                        }

                        if (pVal.ItemUID == "TToDoc")
                        {
                            SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("TToDoc").Specific;
                            var oldVal = coretaxModel.ToDocNum;

                            if (string.IsNullOrWhiteSpace(oEdit.Value))
                            {
                                coretaxModel.ToDocNum = null;
                                coretaxModel.ToDocEntry = null;

                                if (oldVal != null)
                                {
                                    ResetDetail(oForm);
                                }
                            }
                            else if (int.TryParse(oEdit.Value, out int newVal))
                            {
                                if (oldVal == null || newVal != oldVal)
                                {
                                    coretaxModel.ToDocNum = newVal;
                                    ResetDetail(oForm);
                                }
                            }
                        }

                        if (pVal.ItemUID == "TCustFrom")
                        {
                            var oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("TCustFrom").Specific;
                            var newVal = string.IsNullOrWhiteSpace(oEdit.Value) ? null : oEdit.Value;
                            var oldVal = coretaxModel.FromCust;

                            // update model if changed
                            if (!string.Equals(newVal, oldVal, StringComparison.Ordinal))
                            {
                                coretaxModel.FromCust = newVal;
                                ResetDetail(oForm);
                            }
                        }

                        if (pVal.ItemUID == "TCustTo")
                        {
                            var oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("TCustTo").Specific;
                            var newVal = string.IsNullOrWhiteSpace(oEdit.Value) ? null : oEdit.Value;
                            var oldVal = coretaxModel.ToCust;

                            // update model if changed
                            if (!string.Equals(newVal, oldVal, StringComparison.Ordinal))
                            {
                                coretaxModel.ToCust = newVal;
                                ResetDetail(oForm);
                            }
                        }

                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                        SetBoldUnderlinedLabel(oForm, "lDisp");
                        SetBoldUnderlinedLabel(oForm, "lParam");
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MtFind").Specific;
                        oMatrix.AutoResizeColumns();
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE &&
                        pVal.ActionSuccess == true &&
                        pVal.BeforeAction == false)
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                        Task.Run(async() => {
                            await GetNewModel(oForm);
                        }).ContinueWith(task=> MapModelToUI(oForm));
                    }
                    
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction) // after the press
                    {
                        //Check Box
                        if (new[] { "CkInv", "CkDp", "CkCM" }.Contains(pVal.ItemUID))
                        {
                            CheckBoxHandler(pVal, FormUID);
                        }

                        //
                        if (pVal.ItemUID == "BtFind")
                        {
                            BtnFindHandler(pVal, FormUID);
                        }

                        //
                        if (pVal.ItemUID == "BtGen")
                        {
                            BtnGenHandler(pVal, FormUID);
                        }

                        //
                        if (pVal.ItemUID == "BtXML")
                        {
                            ExportToXml(FormUID);
                        }
                        
                        //
                        if (pVal.ItemUID == "BtCSV")
                        {
                            ExportToCsv(FormUID);
                        }

                        //
                        if (pVal.ItemUID == "BtSave")
                        {
                            if (!coretaxModel.IsARCreditMemo && !coretaxModel.IsARDownPayment && !coretaxModel.IsARInvoice)
                                throw new Exception("Please select document to display.");

                            bool isNew = coretaxModel.DocEntry == null || coretaxModel.DocEntry == 0;
                            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

                            // Confirmation dialog
                            int response = Application.SBO_Application.MessageBox(
                                $"Are you sure you want to {(isNew ? "add" : "update")} this document?",
                                1,
                                "Yes",
                                "No",
                                ""
                            );

                            if (response == 1) // Yes
                            {
                                Task.Run(() => {
                                    int docEntry = 0;

                                    try
                                    {
                                        // ✅ Run async service but wait synchronously
                                        if (isNew)
                                            docEntry = TransactionService.AddDataCoretax(coretaxModel).GetAwaiter().GetResult();
                                        else
                                            docEntry = TransactionService.UpdateDataCoretax(coretaxModel).GetAwaiter().GetResult();
                                    }
                                    catch (Exception ex)
                                    {
                                        Application.SBO_Application.StatusBar.SetText($"Error saving: {ex.Message}",
                                            SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        throw;
                                    }

                                    if (docEntry == 0)
                                        throw new Exception($"Failed to {(isNew ? "add" : "update")} document.");

                                    Application.SBO_Application.MessageBox("Successfully saved.");

                                    // ✅ Reload the model synchronously too
                                    try
                                    {
                                        coretaxModel = TransactionService.GetCoretaxByKey(docEntry).GetAwaiter().GetResult();
                                    }
                                    catch (Exception ex)
                                    {
                                        Application.SBO_Application.StatusBar.SetText($"Error reload: {ex.Message}",
                                            SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        throw;
                                    }
                                }).ContinueWith(task => {
                                    // Refresh UI
                                    MapModelToUI(oForm);
                                });
                            }
                        }

                        if (pVal.ItemUID == "BtClose")
                        {
                            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                            // Confirmation dialog
                            int response = Application.SBO_Application.MessageBox(
                                $"Are you sure you want to close this document?",
                                1,
                                "Yes",
                                "No",
                                ""
                            );

                            if (response == 1) // Yes
                            {
                                Task.Run(async() => {
                                    int docEntry = coretaxModel.DocEntry ?? 0;

                                    try
                                    {
                                        await TransactionService.CloseCoretax(docEntry);
                                        await TransactionService.UpdateStatusInv(coretaxModel);
                                    }
                                    catch (Exception ex)
                                    {
                                        Application.SBO_Application.StatusBar.SetText($"Error closing: {ex.Message}",
                                            SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        throw;
                                    }
                                    Application.SBO_Application.MessageBox("Successfully closed.");

                                    // ✅ Reload the model synchronously too
                                    try
                                    {
                                        coretaxModel = TransactionService.GetCoretaxByKey(docEntry).GetAwaiter().GetResult();
                                    }
                                    catch (Exception ex)
                                    {
                                        Application.SBO_Application.StatusBar.SetText($"Error reload: {ex.Message}",
                                            SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        throw;
                                    }
                                }).ContinueWith(task => {
                                    // Refresh UI
                                    MapModelToUI(oForm);
                                });
                            }
                        }

                        if (pVal.ItemUID == "BtCancel")
                        {
                            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                            oForm.Close();
                        }
                    }

                    //if (pVal.ItemUID == "MtFind"
                    //&& pVal.BeforeAction
                    //&& pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    //{
                    //    if (pVal.ColUID != "Col_10")
                    //    {
                    //        BubbleEvent = false; // stop opening editor
                    //        return;
                    //    }
                    //}

                    //if (pVal.ItemUID == "MtDetail"
                    //&& pVal.BeforeAction
                    //&& pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                    //{
                    //    BubbleEvent = false; // stop opening editor
                    //    return;
                    //}

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && !pVal.BeforeAction)
                    {
                        CflHandler(pVal, FormUID);
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "MtFind" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false && pVal.ColUID == "Col_10")
                        {
                            SelectFilterHandler(FormUID, pVal.Row);
                        }
                    }

                    if (pVal.FormTypeEx == "SBOAddonCoreTax.CoretaxForm" &&
                        pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE &&
                        pVal.BeforeAction == false)
                    {
                        AdjustGrids(FormUID);
                    }

                }
                catch (Exception ex)
                {
                    Application.SBO_Application.StatusBar.SetText("Error: " + ex.Message,
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }
        }
        
        private async Task GetNewModel(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);
                if (_pb != null) { _pb.Stop(); _pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Data loading...", 0, false);
                Reset(oForm);
                coretaxModel = new CoretaxModel();
                coretaxModel.DocNum = await TransactionService.GetLastDocNum();
                coretaxModel.SeriesId = TransactionService.GetSeriesIdCoretax();
                coretaxModel.SeriesName = TransactionService.GetSeriesName(coretaxModel.SeriesId ?? 0);
                coretaxModel.DocDate = DateTime.Today;
                coretaxModel.Status = "O";
            }
            catch (Exception e)
            {

                throw e;
            }
            finally
            {
                oForm.Freeze(false);
                if (_pb != null) { _pb.Stop(); _pb = null; }
            }
        }

        private void ExportToXml(string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

            try
            {
                oForm.Freeze(true);
                if (_pb != null) { _pb.Stop(); _pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Exporting data to XML...", 0, false);
                
                if (coretaxModel.Detail != null && coretaxModel.Detail.Any())
                {
                    var listOfTax = new List<TaxInvoice>();

                    foreach (var item in FindListModel)
                    {
                        var docEntry = item.DocEntry;
                        var docNo = item.DocNo;
                        var listDocInv = coretaxModel.Detail.Where((g) => g.DocEntry == docEntry && g.NoDocument == docNo).ToList().OrderBy((c)=>c.LineNum);
                        if (listDocInv != null && listDocInv.Any())
                        {
                            //var invExp = listDocInv.Where((i) => i.Status == "Y").ToList();
                            //if (invExp != null && invExp.Any())
                            //{

                            //}
                            //else
                            //{

                            //}

                            foreach (var itemInv in listDocInv)
                            {
                                var taxInv = new TaxInvoice
                                {
                                    TaxInvoiceDate = itemInv.InvDate,
                                    TaxInvoiceOpt = "Normal",
                                    TrxCode = itemInv.JenisPajak,
                                    AddInfo = itemInv.JenisPajak == "04" ? "" : itemInv.AddInfo,
                                    CustomDoc = "",
                                    CustomDocMonthYear = "",
                                    RefDesc = itemInv.Referensi,
                                    FacilityStamp = "",
                                    SellerIDTKU = itemInv.SellerIDTKU,
                                    BuyerTin = itemInv.NomorNPWP,
                                    BuyerDocument = itemInv.BuyerDocument,
                                    BuyerCountry = itemInv.BuyerCountry,
                                    BuyerDocumentNumber = itemInv.NomorNPWP,
                                    BuyerName = itemInv.BPName,
                                    BuyerAdress = itemInv.NPWPAddress,
                                    BuyerEmail = itemInv.BuyerEmail,
                                    BuyerIDTKU = itemInv.BuyerIDTKU,
                                };
                                listOfTax.Add(taxInv);
                            }
                        }
                    }

                    if (listOfTax.Any())
                    {
                        var TIN = coretaxModel.Detail.First().TIN;
                        var invoice = new TaxInvoiceBulk
                        {
                            TIN = TIN,
                            ListOfTaxInvoice = new ListOfTaxInvoice
                            {
                                TaxInvoiceCollection = listOfTax
                            }
                        };
                        
                        if(ExportService.ExportXml(invoice))
                            Application.SBO_Application.MessageBox("Successfully Exported to XML.");
                    }
                    else
                    {
                        throw new Exception("No data to Export.");
                    }
                }
                else
                {
                    throw new Exception("No data to Export.");
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Error: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                if (_pb != null) { _pb.Stop(); _pb = null; }
            }
        }

        private void ExportToCsv(string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

            try
            {
                oForm.Freeze(true);
                if (_pb != null) { _pb.Stop(); _pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Exporting data to XML...", 0, false);

                if (coretaxModel.Detail != null && coretaxModel.Detail.Any())
                {
                    var listOfTax = new List<TaxInvoice>();

                    foreach (var item in FindListModel)
                    {
                        var docEntry = item.DocEntry;
                        var docNo = item.DocNo;
                        var listDocInv = coretaxModel.Detail.Where((g) => g.DocEntry == docEntry && g.NoDocument == docNo).ToList().OrderBy((c) => c.LineNum);
                        if (listDocInv != null && listDocInv.Any())
                        {
                            foreach (var itemInv in listDocInv)
                            {
                                var taxInv = new TaxInvoice
                                {
                                    TaxInvoiceDate = itemInv.InvDate,
                                    TaxInvoiceOpt = "Normal",
                                    TrxCode = itemInv.JenisPajak,
                                    AddInfo = itemInv.JenisPajak == "04" ? "" : itemInv.AddInfo,
                                    CustomDoc = "",
                                    CustomDocMonthYear = "",
                                    RefDesc = itemInv.Referensi,
                                    FacilityStamp = "",
                                    SellerIDTKU = itemInv.SellerIDTKU,
                                    BuyerTin = itemInv.NomorNPWP,
                                    BuyerDocument = itemInv.BuyerDocument,
                                    BuyerCountry = itemInv.BuyerCountry,
                                    BuyerDocumentNumber = itemInv.NomorNPWP,
                                    BuyerName = itemInv.BPName,
                                    BuyerAdress = itemInv.NPWPAddress,
                                    BuyerEmail = itemInv.BuyerEmail,
                                    BuyerIDTKU = itemInv.BuyerIDTKU,
                                };
                                listOfTax.Add(taxInv);
                            }
                        }
                    }

                    if (listOfTax.Any())
                    {
                        var TIN = coretaxModel.Detail.First().TIN;
                        var invoice = new TaxInvoiceBulk
                        {
                            TIN = TIN,
                            ListOfTaxInvoice = new ListOfTaxInvoice
                            {
                                TaxInvoiceCollection = listOfTax
                            }
                        };

                        if (ExportService.ExportCsv(invoice))
                            Application.SBO_Application.MessageBox("Successfully Exported to XML.");
                    }
                    else
                    {
                        throw new Exception("No data to Export.");
                    }
                }
                else
                {
                    throw new Exception("No data to Export.");
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Error: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                if (_pb != null) { _pb.Stop(); _pb = null; }
            }
        }

        private DateTime? EditTextToDate(string FormUID, string id)
        {
            DateTime? resultDate = null;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            SAPbouiCOM.Item item = oForm.Items.Item(id);
            SAPbouiCOM.EditText editText = (SAPbouiCOM.EditText)item.Specific;
            string inputDate = editText.Value;
            if (DateTime.TryParseExact(
                inputDate,
                "dd/MM/yyyy",          // format used in EditText
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None,
                out DateTime parsedDate
            ))
            {
                resultDate = parsedDate;
            }
            return resultDate;
        }

        private void AdjustGrids(string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

            // Get form width and height
            int formWidth = oForm.ClientWidth;
            int formHeight = oForm.ClientHeight;

            SAPbouiCOM.Item grid1 = oForm.Items.Item("MtFind");

            grid1.Width = formWidth - 20;
            grid1.Height = Convert.ToInt32(formHeight * 0.25);

            SAPbouiCOM.Item mt2 = oForm.Items.Item("MtDetail");
            mt2.Width = formWidth - 20;
        }
        
        private void CheckBoxHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            try
            {
                if(_pb != null) { _pb.Stop(); _pb = null; }
                oForm.Freeze(true);
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("...", 0, false);
                
                if (pVal.ItemUID == "CkInv")
                {
                    SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkInv").Specific;
                    bool isChecked = oCheckBox.Checked;
                    SelectedCkBox["OINV"] = isChecked;
                    coretaxModel.IsARInvoice = isChecked;
                }
                if (pVal.ItemUID == "CkDp")
                {
                    SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkDp").Specific;
                    bool isChecked = oCheckBox.Checked;
                    SelectedCkBox["ODPI"] = isChecked;
                    coretaxModel.IsARDownPayment = isChecked;
                }
                if (pVal.ItemUID == "CkCM")
                {
                    SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkCM").Specific;
                    bool isChecked = oCheckBox.Checked;
                    SelectedCkBox["ORIN"] = isChecked;
                    coretaxModel.IsARCreditMemo = isChecked;
                }

                if (SelectedCkBox.ContainsValue(true))
                {
                    FilterIsShow = true;
                }
                else
                {
                    FilterIsShow = false;
                }

                ShowFilterGroup(oForm);
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                oForm.Freeze(false);
                if (_pb != null) { _pb.Stop(); _pb = null; }
            }
        }

        private void GetDataComboBox(SAPbouiCOM.Form form, string id, string table)
        {
            SAPbouiCOM.Item comboItem = form.Items.Item(id);
            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)comboItem.Specific;

            // Remove any default values
            while (oCombo.ValidValues.Count > 0)
            {
                oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            // Add dropdown values
            if (table == "OUBR")
            {
                if (cbBranchValues.Any())
                {
                    foreach (var item in cbBranchValues)
                    {
                        oCombo.ValidValues.Add(item.Key, item.Value);
                    }
                }
                else
                {
                    cbBranchValues = GetDataCbBranch();
                    foreach (var item in cbBranchValues)
                    {
                        oCombo.ValidValues.Add(item.Key, item.Value);
                    }
                }
            }
            if (table == "OOCR")
            {
                if (cbOutletValues.Any())
                {
                    foreach (var item in cbOutletValues)
                    {
                        oCombo.ValidValues.Add(item.Key, item.Value);
                    }
                }
                else
                {
                    cbOutletValues = GetDataCbOutlet();
                    foreach (var item in cbOutletValues)
                    {
                        oCombo.ValidValues.Add(item.Key, item.Value);
                    }
                }
                //// 3. Load values using DI API Recordset
                //Company oCompany = Services.CompanyService.GetCompany();
                //SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //rs.DoQuery($"SELECT Code, Name FROM {table} ORDER BY Code"); // replace with your query

                //while (!rs.EoF)
                //{
                //    string code = rs.Fields.Item("Code").Value.ToString();
                //    string name = rs.Fields.Item("Name").Value.ToString();

                //    oCombo.ValidValues.Add(code, name);
                //    rs.MoveNext();
                //}

            }
        }

        private Dictionary<string,string> GetDataCbBranch()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            try
            {
                Company oCompany = Services.CompanyService.GetCompany();
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery($"SELECT Code, Name FROM OUBR ORDER BY Code"); // replace with your query

                while (!rs.EoF)
                {
                    string code = rs.Fields.Item("Code").Value.ToString();
                    string name = rs.Fields.Item("Name").Value.ToString();

                    result.Add(code, name);
                    rs.MoveNext();
                }
                return result;
            }
            catch (Exception)
            {

                throw;
            }
        }

        private Dictionary<string, string> GetDataCbOutlet()
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            try
            {
                Company oCompany = Services.CompanyService.GetCompany();
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery($"SELECT T0.[OcrCode] AS Code, T0.[OcrName] AS Name FROM OOCR T0 WHERE T0.[DimCode] = '4' "); // replace with your query

                while (!rs.EoF)
                {
                    string code = rs.Fields.Item("Code").Value.ToString();
                    string name = rs.Fields.Item("Name").Value.ToString();

                    result.Add(code, name);
                    rs.MoveNext();
                }
                return result;
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void CflHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.IChooseFromListEvent oCFLEvent = (SAPbouiCOM.IChooseFromListEvent)pVal;
            if (SelectedCkBox.Where((ck)=>ck.Value).Count() == 1)
            {
                var selectedCk = SelectedCkBox.Where((ck) => ck.Value).First().Key;
                if (oCFLEvent.ChooseFromListUID == "CflDocFrom" + selectedCk)
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvent.SelectedObjects;

                    if (oDataTable != null && oDataTable.Rows.Count > 0)
                    {
                        // Get values from the selected row
                        string docNum = oDataTable.GetValue("DocNum", 0).ToString();
                        string docEntry = oDataTable.GetValue("DocEntry", 0).ToString();
                        if (coretaxModel.FromDocEntry != int.Parse(docEntry))
                        {
                            //coretaxModel.FromDocNum = int.Parse(docNum);
                            coretaxModel.FromDocEntry = int.Parse(docEntry);

                            // Get the form object
                            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

                            // Set values into your form fields (bound to User Data Sources)
                            oForm.DataSources.UserDataSources.Item("DocFromDS").ValueEx = docNum;
                        }
                    }
                }
                if (oCFLEvent.ChooseFromListUID == "CflDocTo" + selectedCk)
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvent.SelectedObjects;

                    if (oDataTable != null && oDataTable.Rows.Count > 0)
                    {
                        // Get values from the selected row
                        string docNum = oDataTable.GetValue("DocNum", 0).ToString();
                        string docEntry = oDataTable.GetValue("DocEntry", 0).ToString();

                        if (coretaxModel.ToDocEntry != int.Parse(docEntry))
                        {
                            //coretaxModel.ToDocNum = int.Parse(docNum);
                            coretaxModel.ToDocEntry = int.Parse(docEntry);

                            // Get the form object
                            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

                            // Set values into your form fields (bound to User Data Sources)
                            oForm.DataSources.UserDataSources.Item("DocToDS").ValueEx = docNum;
                        }
                    }
                }
            }
            
            if (oCFLEvent.ChooseFromListUID == "CflCustFrom")
            {
                SAPbouiCOM.DataTable oDataTable = oCFLEvent.SelectedObjects;

                if (oDataTable != null && oDataTable.Rows.Count > 0)
                {
                    // Get values from the selected row
                    string code = oDataTable.GetValue("CardCode", 0).ToString();
                    if (code != coretaxModel.FromCust)
                    {
                        //coretaxModel.FromCust = code;
                        // Get the form object
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

                        // Set values into your form fields (bound to User Data Sources)
                        oForm.DataSources.UserDataSources.Item("CustFromDS").ValueEx = code;
                    }
                }
            }
            if (oCFLEvent.ChooseFromListUID == "CflCustTo")
            {
                SAPbouiCOM.DataTable oDataTable = oCFLEvent.SelectedObjects;

                if (oDataTable != null && oDataTable.Rows.Count > 0)
                {
                    // Get values from the selected row
                    string code = oDataTable.GetValue("CardCode", 0).ToString();
                    if (coretaxModel.ToCust != code)
                    {
                        //coretaxModel.ToCust = code;

                        // Get the form object
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

                        // Set values into your form fields (bound to User Data Sources)
                        oForm.DataSources.UserDataSources.Item("CustToDS").ValueEx = code;
                    }
                }
            }
        }

        private void BtnFindHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            Task.Run(async () => {
                try
            {
                oForm.Freeze(true);
                if (_pb != null) { _pb.Stop(); _pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Retrieving data...", 0, false);
                _pb.Text = "Retrieving data...";
                FindListModel.Clear();
                coretaxModel.Detail.Clear();
                ClearMatrix(oForm, "MtDetail", "DT_DETAIL", "@T2_CORETAX_DT");

                bool allDt = false;
                bool allDoc = false;
                bool allCust = false;
                bool allBranch = false;
                bool allOutlet = false;
                string dtFrom = string.Empty;
                string dtTo = string.Empty;
                string docFrom = string.Empty;
                string docTo = string.Empty;
                string custFrom = string.Empty;
                string custTo = string.Empty;
                string branchFrom = string.Empty;
                string branchTo = string.Empty;
                string outFrom = string.Empty;
                string outTo = string.Empty;

                if (oForm.Items.Item("CkAllDt").Enabled)
                //if (ItemIsExists(oForm, "CkAllDt") && oForm.Items.Item("CkAllDt").Enabled)
                {
                    SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllDt").Specific;
                    allDt = oCk.Checked;
                }
                if (oForm.Items.Item("CkAllDoc").Enabled)
                //if (ItemIsExists(oForm, "CkAllDoc") && oForm.Items.Item("CkAllDoc").Enabled)
                {
                    SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllDoc").Specific;
                    allDoc = oCk.Checked;
                }
                if (oForm.Items.Item("CkAllCust").Enabled)
                //if (ItemIsExists(oForm, "CkAllCust") && oForm.Items.Item("CkAllCust").Enabled)
                {
                    SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllCust").Specific;
                    allCust = oCk.Checked;
                }
                if (oForm.Items.Item("CkAllBr").Enabled)
                //if (ItemIsExists(oForm, "CkAllBr") && oForm.Items.Item("CkAllBr").Enabled)
                {
                    SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllBr").Specific;
                    allBranch = oCk.Checked;
                }
                if (oForm.Items.Item("CkAllOtl").Enabled)
                //if (ItemIsExists(oForm, "CkAllOtl") && oForm.Items.Item("CkAllOtl").Enabled)
                {
                    SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllOtl").Specific;
                    allOutlet = oCk.Checked;
                }

                if (!allDt)
                {
                    if (oForm.Items.Item("TFromDt").Enabled)
                    //if (ItemIsExists(oForm, "TFromDt") && oForm.Items.Item("TFromDt").Enabled)
                    {
                        SAPbouiCOM.EditText oDtFrom = (SAPbouiCOM.EditText)oForm.Items.Item("TFromDt").Specific;
                        if (DateTime.TryParseExact(oDtFrom.Value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDtFrom))
                        {
                            dtFrom = parsedDtFrom.ToString("yyyy-MM-dd");
                        }
                    }
                    if (oForm.Items.Item("TToDt").Enabled)
                    //if (ItemIsExists(oForm, "TToDt") && oForm.Items.Item("TToDt").Enabled)
                    {
                        SAPbouiCOM.EditText oDtTo = (SAPbouiCOM.EditText)oForm.Items.Item("TToDt").Specific;
                        if (DateTime.TryParseExact(oDtTo.Value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDtTo))
                        {
                            dtTo = parsedDtTo.ToString("yyyy-MM-dd");
                        }
                    }
                }
                if (!allDoc)
                {
                    if (oForm.Items.Item("TFromDoc").Enabled)
                    //if (ItemIsExists(oForm, "TFromDoc") && oForm.Items.Item("TFromDoc").Enabled)
                    {
                        docFrom = coretaxModel.FromDocEntry != 0 ? coretaxModel.FromDocEntry.ToString() : string.Empty;
                        //SAPbouiCOM.EditText oDocFrom = (SAPbouiCOM.EditText)oForm.Items.Item("TFromDoc").Specific;
                        //if (!string.IsNullOrEmpty(oDocFrom.Value.Trim()))
                        //{
                        //    docFrom = oDocFrom.Value.Trim();
                        //}
                    }
                    if (oForm.Items.Item("TToDoc").Enabled)
                    //if (ItemIsExists(oForm, "TToDoc") && oForm.Items.Item("TToDoc").Enabled)
                    {
                        docTo = coretaxModel.ToDocEntry != 0 ? coretaxModel.ToDocEntry.ToString() : string.Empty;
                        //SAPbouiCOM.EditText oDocTo = (SAPbouiCOM.EditText)oForm.Items.Item("TToDoc").Specific;
                        //if (!string.IsNullOrEmpty(oDocTo.Value.Trim()))
                        //{
                        //    docTo = oDocTo.Value.Trim();
                        //}
                    }
                }
                if (!allCust)
                {
                    if (oForm.Items.Item("TCustFrom").Enabled)
                    //if (ItemIsExists(oForm, "TCustFrom") && oForm.Items.Item("TCustFrom").Enabled)
                    {
                        SAPbouiCOM.EditText oCustFrom = (SAPbouiCOM.EditText)oForm.Items.Item("TCustFrom").Specific;
                        if (!string.IsNullOrEmpty(oCustFrom.Value.Trim()))
                        {
                            custFrom = oCustFrom.Value.Trim();
                        }
                    }
                    if (oForm.Items.Item("TCustTo").Enabled)
                    //if (ItemIsExists(oForm, "TCustTo") && oForm.Items.Item("TCustTo").Enabled)
                    {
                        SAPbouiCOM.EditText oCustTo = (SAPbouiCOM.EditText)oForm.Items.Item("TCustTo").Specific;
                        if (!string.IsNullOrEmpty(oCustTo.Value.Trim()))
                        {
                            custTo = oCustTo.Value.Trim();
                        }
                    }
                }
                if (!allBranch)
                {
                    if (oForm.Items.Item("CbBrFrom").Enabled)
                    //if (ItemIsExists(oForm, "CbBrFrom") && oForm.Items.Item("CbBrFrom").Enabled)
                    {
                        SAPbouiCOM.ComboBox oBrFrom = (SAPbouiCOM.ComboBox)oForm.Items.Item("CbBrFrom").Specific;
                        if (!string.IsNullOrEmpty(oBrFrom.Value.Trim()))
                        {
                            branchFrom = oBrFrom.Value.Trim();
                        }
                    }
                    if (oForm.Items.Item("CbBrTo").Enabled)
                    //if (ItemIsExists(oForm, "CbBrTo") && oForm.Items.Item("CbBrTo").Enabled)
                    {
                        SAPbouiCOM.ComboBox oBrTo = (SAPbouiCOM.ComboBox)oForm.Items.Item("CbBrTo").Specific;
                        if (!string.IsNullOrEmpty(oBrTo.Value.Trim()))
                        {
                            branchTo = oBrTo.Value.Trim();
                        }
                    }
                }
                if (!allOutlet)
                {
                    if (oForm.Items.Item("CbOtlFrom").Enabled)
                    //if (ItemIsExists(oForm, "CbOtlFrom") && oForm.Items.Item("CbOtlFrom").Enabled)
                    {
                        SAPbouiCOM.ComboBox oOtlFrom = (SAPbouiCOM.ComboBox)oForm.Items.Item("CbOtlFrom").Specific;
                        if (!string.IsNullOrEmpty(oOtlFrom.Value.Trim()))
                        {
                            outFrom = oOtlFrom.Value.Trim();
                        }
                    }
                    if (oForm.Items.Item("CbOtlTo").Enabled)
                    //if (ItemIsExists(oForm, "CbOtlTo") && oForm.Items.Item("CbOtlTo").Enabled)
                    {
                        SAPbouiCOM.ComboBox oOtlTo = (SAPbouiCOM.ComboBox)oForm.Items.Item("CbOtlTo").Specific;
                        if (!string.IsNullOrEmpty(oOtlTo.Value.Trim()))
                        {
                            outTo = oOtlTo.Value.Trim();
                        }
                    }
                }
                FindListModel = await TransactionService.GetDataFilter(
                            SelectedCkBox, dtFrom, dtTo, docFrom, docTo, custFrom, custTo,
                            branchFrom, branchTo, outFrom, outTo
                            );
                    SetMtFind(oForm);
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                oForm.Freeze(false);
                if (_pb != null) { _pb.Stop(); _pb = null; }
            }
            });
        }
        
        private void SetMtFind(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MtFind").Specific;
                // Create DataTable if not exists
                SAPbouiCOM.DataTable oDT;
                if (!DtIsExists(oForm, "DT_FILTER"))
                {
                    oDT = oForm.DataSources.DataTables.Add("DT_FILTER");
                }
                else
                {
                    oDT = oForm.DataSources.DataTables.Item("DT_FILTER");
                }
                // Clear previous rows
                oDT.Clear();

                // Also clear matrix (important)
                oMatrix.Clear();

                // Define all columns (make sure sizes are large enough)
                oDT.Columns.Add("Select", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 1);
                oDT.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Integer);
                oDT.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_Integer);
                oDT.Columns.Add("ObjType", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("ObjName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                oDT.Columns.Add("BPCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("BPName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                oDT.Columns.Add("PostDate", SAPbouiCOM.BoFieldsType.ft_Date);
                oDT.Columns.Add("BranchCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("BranchName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                oDT.Columns.Add("OutletCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("OutletName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                oDT.Columns.Add("#", SAPbouiCOM.BoFieldsType.ft_Integer);


                // Fill DataTable from model
                for (int i = 0; i < FindListModel.Count; i++)
                {
                    var row = FindListModel[i];
                    oDT.Rows.Add();

                    // Convert date string to SAP date format
                    if (!string.IsNullOrEmpty(row.PostDate) && DateTime.TryParse(row.PostDate, out var invDate))
                        oDT.SetValue("PostDate", i, invDate.ToString("yyyyMMdd"));
                    else
                        oDT.SetValue("PostDate", i, "");

                    oDT.SetValue("Select", i, row.Selected ? "Y" : "N");
                    oDT.SetValue("DocEntry", i, row.DocEntry ?? "");
                    oDT.SetValue("DocNum", i, row.DocNo ?? "");
                    oDT.SetValue("BPCode", i, row.CardCode ?? "");
                    oDT.SetValue("BPName", i, row.CardName ?? "");
                    oDT.SetValue("ObjType", i, row.ObjType ?? "");
                    oDT.SetValue("ObjName", i, row.ObjName ?? "");
                    oDT.SetValue("BranchCode", i, row.BranchCode ?? "");
                    oDT.SetValue("BranchName", i, row.BranchName ?? "");
                    oDT.SetValue("OutletCode", i, row.OutletCode ?? "");
                    oDT.SetValue("OutletName", i, row.OutletName ?? "");
                    oDT.SetValue("#", i, (i + 1));
                }

                oMatrix.Columns.Item("Col_1").DataBind.Bind("DT_FILTER", "DocNum");
                oMatrix.Columns.Item("Col_1").Width = 80;
                oMatrix.Columns.Item("Col_2").DataBind.Bind("DT_FILTER", "BPCode");
                oMatrix.Columns.Item("Col_2").Width = 80;
                oMatrix.Columns.Item("Col_3").DataBind.Bind("DT_FILTER", "BPName");
                oMatrix.Columns.Item("Col_3").Width = 100;
                oMatrix.Columns.Item("Col_4").DataBind.Bind("DT_FILTER", "ObjName");
                oMatrix.Columns.Item("Col_4").Width = 100;
                oMatrix.Columns.Item("Col_5").DataBind.Bind("DT_FILTER", "PostDate");
                oMatrix.Columns.Item("Col_5").Width = 80;
                oMatrix.Columns.Item("Col_6").DataBind.Bind("DT_FILTER", "BranchCode");
                oMatrix.Columns.Item("Col_6").Width = 80;
                oMatrix.Columns.Item("Col_7").DataBind.Bind("DT_FILTER", "BranchName");
                oMatrix.Columns.Item("Col_7").Width = 100;
                oMatrix.Columns.Item("Col_8").DataBind.Bind("DT_FILTER", "OutletCode");
                oMatrix.Columns.Item("Col_8").Width = 80;
                oMatrix.Columns.Item("Col_9").DataBind.Bind("DT_FILTER", "OutletName");
                oMatrix.Columns.Item("Col_9").Width = 100;
                oMatrix.Columns.Item("Col_10").DataBind.Bind("DT_FILTER", "Select");
                oMatrix.Columns.Item("Col_10").Width = 40;
                oMatrix.Columns.Item("#").DataBind.Bind("DT_FILTER", "#");
                oMatrix.Columns.Item("#").Width = 30;
                
                // Load data into matrix
                oMatrix.LoadFromDataSource();
                oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                int white = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                for (int i = 0; i < oMatrix.RowCount; i++)
                {
                    oMatrix.CommonSetting.SetRowBackColor(i + 1, white);
                }
                oMatrix.AutoResizeColumns();
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        private void SelectFilterHandler(string FormUID, int Row)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            try
            {
                if (FindListModel.Any())
                {
                    if (Row == 0)
                    {
                        ToggleSelectAll(oForm);
                    }
                    else
                    {
                        ToggleSelectSingle(oForm, Row);
                    }
                    SAPbouiCOM.Item btGen = oForm.Items.Item("BtGen");
                    if (FindListModel != null && FindListModel.Any((f) => f.Selected))
                    {
                        btGen.Enabled = true;
                    }
                    else
                    {
                        btGen.Enabled = false;
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                ClearMatrix(oForm, "MtDetail", "DT_DETAIL", "@T2_CORETAX_DT");
            }
        }

        private void ToggleSelectAll(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);
                if (_pb != null) { _pb.Stop(); _pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Selecting data...", 0, false);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MtFind").Specific;

                bool selectAll = oMatrix.Columns.Item("Col_10").TitleObject.Caption != "[X]";
                
                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Col_10").Cells.Item(i).Specific).Checked = selectAll;
                }
                oMatrix.Columns.Item("Col_10").TitleObject.Caption = selectAll ? "[X]" : "[ ]";
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Error: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                oForm.Freeze(false);
                if (_pb != null) { _pb.Stop(); _pb = null; }
            }
        }

        private void ToggleSelectSingle(SAPbouiCOM.Form oForm, int mtRow)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MtFind").Specific;

                // Get clicked row
                int row = mtRow - 1;

                // Get checkbox value (grid stores it as string "Y"/"N" or "tYES"/"tNO")
                bool isChecked = ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Col_10").Cells.Item(mtRow).Specific).Checked;
                SAPbouiCOM.DataTable oDT;
                if (!DtIsExists(oForm, "DT_FILTER"))
                {
                    oDT = oForm.DataSources.DataTables.Add("DT_FILTER");
                }
                else
                {
                    oDT = oForm.DataSources.DataTables.Item("DT_FILTER");
                }
                string docEntryVal = oDT.GetValue("DocEntry",row).ToString();
                string objTypeVal = oDT.GetValue("ObjType", row).ToString();
                var tempData = FindListModel.Where((f) => f.DocEntry == docEntryVal && f.ObjType == objTypeVal).FirstOrDefault();
                if (tempData != null)
                {
                    tempData.Selected = isChecked;
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void BtnGenHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            Task.Run(async () =>
            {
                try
            {
                oForm.Freeze(true);
                if (_pb != null) { _pb.Stop(); _pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Generating data...", 0, false);

                coretaxModel.Detail.Clear();
                var filteredHeader = FindListModel.Where((f) => f.Selected).ToList();
                if (filteredHeader != null && filteredHeader.Any())
                {
                    coretaxModel.Detail = await TransactionService.GetDataGenerate(filteredHeader);
                    SetMtGenerateByModel(oForm);
                }
            }
            catch (Exception e)
            {

                throw e;
            }
            finally
            {
                oForm.Freeze(false);
                    if (_pb != null) { _pb.Stop(); _pb = null; }
            }
            });
        }

        private void SetBtnExport(SAPbouiCOM.Form oForm)
        {
            if (coretaxModel.Detail.Any() && coretaxModel.DocEntry != null && coretaxModel.DocEntry != 0)
            {
                oForm.Items.Item("BtCSV").Enabled = true;
                oForm.Items.Item("BtXML").Enabled = true;
            }
            else
            {
                oForm.Items.Item("BtCSV").Enabled = false;
                oForm.Items.Item("BtXML").Enabled = false;
            }
        }

        private void RemoveFocus(SAPbouiCOM.Form oForm)
        {
            if (!ItemIsExists(oForm, "dummy"))
            {
                var dummyItem = oForm.Items.Add("dummy", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                dummyItem.Left = 5;
                dummyItem.Top = 5;
                dummyItem.Width = 1;
                dummyItem.Height = 1;
                ((SAPbouiCOM.EditText)dummyItem.Specific).Value = "";
            }

            // Force focus to dummy edit field
            oForm.Items.Item("dummy").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            oForm.ActiveItem = "dummy";
        }

        private void ShowFilterGroup(SAPbouiCOM.Form oForm)
        {
            RemoveFocus(oForm);
            try
            {
                string[] dateItems = { "TFromDt", "TToDt", "CkAllDt" };
                string[] docItems = { "TFromDoc", "TToDoc", "CkAllDoc" };
                string[] custItems = { "TCustFrom", "TCustTo", "CkAllCust" };
                string[] brItems = { "CbBrFrom", "CbBrTo", "CkAllBr" };
                string[] otlItems = { "CbOtlFrom", "CbOtlTo", "CkAllOtl" };

                if (FilterIsShow)
                {
                    int countDoc = SelectedCkBox.Count(d => d.Value);

                    // Always enable date filter
                    SetEnabled(oForm, dateItems, true);

                    // Enable doc filter only if exactly 1 doc type selected
                    SetEnabled(oForm, docItems, countDoc == 1);

                    if (countDoc == 1)
                    {
                        var table = SelectedCkBox.First(d => d.Value).Key;
                        SetDocumentCfl(oForm, "CflDocFrom" + table, "DocFromDS", "TFromDoc");
                        SetDocumentCfl(oForm, "CflDocTo" + table, "DocToDS", "TToDoc");
                    }

                    // Always enable customer, branch, outlet filters
                    SetEnabled(oForm, custItems, true);
                    SetEnabled(oForm, brItems, true);
                    SetEnabled(oForm, otlItems, true);

                    // Set CFLs
                    SetCustomerCfl(oForm, "CflCustFrom", "CustFromDS", "TCustFrom");
                    SetCustomerCfl(oForm, "CflCustTo", "CustToDS", "TCustTo");

                    // Fill combos
                    GetDataComboBox(oForm, "CbBrFrom", "OUBR");
                    GetDataComboBox(oForm, "CbBrTo", "OUBR");
                    GetDataComboBox(oForm, "CbOtlFrom", "OOCR");
                    GetDataComboBox(oForm, "CbOtlTo", "OOCR");

                    oForm.Items.Item("BtFind").Enabled = true;
                }
                else
                {
                    // Hide/disable everything if FilterIsShow = false
                    SetEnabled(oForm, dateItems, false);
                    SetEnabled(oForm, docItems, false);
                    SetEnabled(oForm, custItems, false);
                    SetEnabled(oForm, brItems, false);
                    SetEnabled(oForm, otlItems, false);

                    ClearCombo(oForm, "CbBrFrom");
                    ClearCombo(oForm, "CbBrTo");
                    ClearCheckBox(oForm, "CkAllBr", "CkBrDS");

                    ClearCombo(oForm, "CbOtlFrom");
                    ClearCombo(oForm, "CbOtlTo");
                    ClearCheckBox(oForm, "CkAllOtl", "CkOtlDS");

                    oForm.Items.Item("BtFind").Enabled = false;
                }

                // Always clear filter fields
                ClearEdit(oForm, "TFromDt");
                ClearEdit(oForm, "TToDt");
                ClearCheckBox(oForm, "CkAllDt", "CkDtDS");

                ClearEdit(oForm, "TFromDoc");
                ClearEdit(oForm, "TToDoc");
                ClearCheckBox(oForm, "CkAllDoc", "CkDocDS");

                ClearEdit(oForm, "TCustFrom");
                ClearEdit(oForm, "TCustTo");
                ClearCheckBox(oForm, "CkAllCust", "CkCustDS");

                ResetDetail(oForm);
            }
            finally
            {
                RemoveFocus(oForm);
            }
        }

        //private void ShowFilterGroup(SAPbouiCOM.Form oForm)
        //{
        //    RemoveFocus(oForm);
        //    if (FilterIsShow)
        //    {
        //        var countDoc = SelectedCkBox.Where((d) => d.Value).Count();
        //        if (!oForm.Items.Item("TFromDt").Enabled) oForm.Items.Item("TFromDt").Enabled = true;
        //        if (!oForm.Items.Item("TToDt").Enabled) oForm.Items.Item("TToDt").Enabled = true;
        //        if (!oForm.Items.Item("CkAllDt").Enabled) oForm.Items.Item("CkAllDt").Enabled = true;

        //        if (countDoc == 1)
        //        {
        //            var table = SelectedCkBox.Where((d) => d.Value).First().Key;
        //            if (!oForm.Items.Item("TFromDoc").Enabled) oForm.Items.Item("TFromDoc").Enabled = true;
        //            if (!oForm.Items.Item("TToDoc").Enabled) oForm.Items.Item("TToDoc").Enabled = true;
        //            if (!oForm.Items.Item("CkAllDoc").Enabled) oForm.Items.Item("CkAllDoc").Enabled = true;
        //            SetDocumentCfl(oForm, "CflDocFrom"+table, "DocFromDS", "TFromDoc");
        //            SetDocumentCfl(oForm, "CflDocTo"+table, "DocToDS", "TToDoc");
        //        }
        //        else
        //        {
        //            if (oForm.Items.Item("TFromDoc").Enabled) oForm.Items.Item("TFromDoc").Enabled = false;
        //            if (oForm.Items.Item("TToDoc").Enabled) oForm.Items.Item("TToDoc").Enabled = false;
        //            if (oForm.Items.Item("CkAllDoc").Enabled) oForm.Items.Item("CkAllDoc").Enabled = false;
        //        }

        //        if (!oForm.Items.Item("TCustFrom").Enabled) oForm.Items.Item("TCustFrom").Enabled = true;
        //        if (!oForm.Items.Item("TCustTo").Enabled) oForm.Items.Item("TCustTo").Enabled = true;
        //        if (!oForm.Items.Item("CkAllCust").Enabled) oForm.Items.Item("CkAllCust").Enabled = true;
        //        if (!oForm.Items.Item("CbBrFrom").Enabled) oForm.Items.Item("CbBrFrom").Enabled = true;
        //        if (!oForm.Items.Item("CbBrTo").Enabled) oForm.Items.Item("CbBrTo").Enabled = true;
        //        if (!oForm.Items.Item("CkAllBr").Enabled) oForm.Items.Item("CkAllBr").Enabled = true;
        //        if (!oForm.Items.Item("CbOtlFrom").Enabled) oForm.Items.Item("CbOtlFrom").Enabled = true;
        //        if (!oForm.Items.Item("CbOtlTo").Enabled) oForm.Items.Item("CbOtlTo").Enabled = true;
        //        if (!oForm.Items.Item("CkAllOtl").Enabled) oForm.Items.Item("CkAllOtl").Enabled = true;

        //        SetCustomerCfl(oForm, "CflCustFrom", "CustFromDS", "TCustFrom");
        //        SetCustomerCfl(oForm, "CflCustTo", "CustToDS", "TCustTo");

        //        GetDataComboBox(oForm, "CbBrFrom", "OUBR");
        //        GetDataComboBox(oForm, "CbBrTo", "OUBR");

        //        GetDataComboBox(oForm, "CbOtlFrom", "OUDP");
        //        GetDataComboBox(oForm, "CbOtlTo", "OUDP");

        //        oForm.Items.Item("BtFind").Enabled = true;
        //    }
        //    else
        //    {
        //        if (oForm.Items.Item("TFromDt").Enabled) oForm.Items.Item("TFromDt").Enabled = false;
        //        if (oForm.Items.Item("TToDt").Enabled) oForm.Items.Item("TToDt").Enabled = false;
        //        if (oForm.Items.Item("CkAllDt").Enabled) oForm.Items.Item("CkAllDt").Enabled = false;
        //        if (oForm.Items.Item("TFromDoc").Enabled) oForm.Items.Item("TFromDoc").Enabled = false;
        //        if (oForm.Items.Item("TToDoc").Enabled) oForm.Items.Item("TToDoc").Enabled = false;
        //        if (oForm.Items.Item("CkAllDoc").Enabled) oForm.Items.Item("CkAllDoc").Enabled = false;
        //        if (oForm.Items.Item("TCustFrom").Enabled) oForm.Items.Item("TCustFrom").Enabled = false;
        //        if (oForm.Items.Item("TCustTo").Enabled) oForm.Items.Item("TCustTo").Enabled = false;
        //        if (oForm.Items.Item("CkAllCust").Enabled) oForm.Items.Item("CkAllCust").Enabled = false;
        //        if (oForm.Items.Item("CbBrFrom").Enabled) oForm.Items.Item("CbBrFrom").Enabled = false;
        //        if (oForm.Items.Item("CbBrTo").Enabled) oForm.Items.Item("CbBrTo").Enabled = false; 
        //        if (oForm.Items.Item("CkAllBr").Enabled) oForm.Items.Item("CkAllBr").Enabled = false;
        //        if (oForm.Items.Item("CbOtlFrom").Enabled) oForm.Items.Item("CbOtlFrom").Enabled = false;
        //        if (oForm.Items.Item("CbOtlTo").Enabled) oForm.Items.Item("CbOtlTo").Enabled = false;
        //        if (oForm.Items.Item("CkAllOtl").Enabled) oForm.Items.Item("CkAllOtl").Enabled = false;

        //        oForm.Items.Item("BtFind").Enabled = false;
        //    }
        //    ClearEdit(oForm, "TFromDt");
        //    ClearEdit(oForm, "TToDt");
        //    ClearCheckBox(oForm, "CkAllDt", "CkDtDS");
        //    ClearEdit(oForm, "TFromDoc");
        //    ClearEdit(oForm, "TToDoc");
        //    ClearCheckBox(oForm, "CkAllDoc", "CkDocDS");
        //    ClearEdit(oForm, "TCustFrom");
        //    ClearEdit(oForm, "TCustTo");
        //    ClearCheckBox(oForm, "CkAllCust", "CkCustDS");
        //    ClearCombo(oForm, "CbBrFrom");
        //    ClearCombo(oForm, "CbBrTo");
        //    ClearCheckBox(oForm, "CkAllBr", "CkBrDS");
        //    ClearCombo(oForm, "CbOtlFrom");
        //    ClearCombo(oForm, "CbOtlTo");
        //    ClearCheckBox(oForm, "CkAllOtl", "CkOtlDS");
        //    ResetDetail(oForm);
        //}

        private void ClearCombo(SAPbouiCOM.Form form, string id)
        {
            //if (!ItemIsExists(form, id)) return;
            SAPbouiCOM.Item comboItem = form.Items.Item(id);
            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)comboItem.Specific;

            // Remove any default values
            while (oCombo.ValidValues.Count > 0)
            {
                oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            oCombo.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
        }

        private void ClearEdit(SAPbouiCOM.Form oForm, string id)
        {
            //if (!ItemIsExists(oForm, id)) return;
            SAPbouiCOM.EditText edit = (SAPbouiCOM.EditText)oForm.Items.Item(id).Specific;
            edit.Value = "";
        }

        private void ClearEditDate(SAPbouiCOM.Form oForm, string id, string ds)
        {
            //if (!ItemIsExists(oForm, id)) return;
            ClearEdit(oForm, id);
            oForm.DataSources.UserDataSources.Item(ds).Value = "00000000"; // -> tampil 30.12.1899
        }

        private void ClearCheckBox(SAPbouiCOM.Form oForm, string id, string ds)
        {
            //if (!ItemIsExists(oForm, id)) return;
            SAPbouiCOM.CheckBox ck = (SAPbouiCOM.CheckBox)oForm.Items.Item(id).Specific;
            ck.Checked = false;
            oForm.DataSources.UserDataSources.Item(ds).Value = "N";
        }

        private void SetBoldUnderlinedLabel(SAPbouiCOM.Form oForm, string id)
        {
            // Get StaticText item
            SAPbouiCOM.StaticText label = (SAPbouiCOM.StaticText)oForm.Items.Item(id).Specific;

            // Bold Underline
            label.Item.TextStyle = (int)(SAPbouiCOM.BoFontStyle.fs_Bold | SAPbouiCOM.BoFontStyle.fs_Underline);
            //if (ItemIsExists(oForm, id))
            //{
            //    // Get StaticText item
            //    SAPbouiCOM.StaticText label = (SAPbouiCOM.StaticText)oForm.Items.Item(id).Specific;

            //    // Bold Underline
            //    label.Item.TextStyle = (int)(SAPbouiCOM.BoFontStyle.fs_Bold | SAPbouiCOM.BoFontStyle.fs_Underline);
            //}
        }

        private void ClearGrid(SAPbouiCOM.Form oForm, string gridItemUID, string dataTableUID)
        {
            try
            {
                // Check if the DataTable exists
                SAPbouiCOM.DataTable dt = null;
                if (DtIsExists(oForm, dataTableUID))
                {
                    dt = oForm.DataSources.DataTables.Item(dataTableUID);
                    dt.Clear();
                }
                else
                {
                    // Recreate DataTable with all columns if it doesn't exist
                    dt = oForm.DataSources.DataTables.Add(dataTableUID);
                    dt.Clear();
                }

                // Bind DataTable to Grid
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(gridItemUID).Specific;
                oGrid.DataTable = dt;

            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage("Error clearing grid: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private void ClearMatrix(SAPbouiCOM.Form oForm, string mtUID, string dtUID = null, string tbUID = null)
        {
            try
            {
                //if (!ItemIsExists(oForm, mtUID)) return;
                oForm.Freeze(true);
                SAPbouiCOM.Matrix oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item(mtUID).Specific;
                if (oMtx.RowCount <= 0) return;
                oMtx.Clear();
                if (!string.IsNullOrEmpty(dtUID))
                {
                    // Check if the DataTable exists
                    SAPbouiCOM.DataTable dt = null;
                    if (DtIsExists(oForm, dtUID))
                    {
                        dt = oForm.DataSources.DataTables.Item(dtUID);
                        dt.Clear();
                    }
                    else
                    {
                        // Recreate DataTable with all columns if it doesn't exist
                        dt = oForm.DataSources.DataTables.Add(dtUID);
                        dt.Clear();
                    }
                }

                if (!string.IsNullOrEmpty(tbUID))
                {
                    SAPbouiCOM.DBDataSource db = null;
                    if (DBSIsExists(oForm, tbUID))
                    {
                        db = oForm.DataSources.DBDataSources.Item(tbUID);
                        db.Clear();
                    }
                    else
                    {
                        // Recreate DataTable with all columns if it doesn't exist
                        db = oForm.DataSources.DBDataSources.Add(tbUID);
                        db.Clear();
                    }
                }
                

            }
            catch (Exception ex)
            {
                Application.SBO_Application.SetStatusBarMessage("Error clearing Matrix: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void SetCustomerCfl(SAPbouiCOM.Form oForm, string id, string ds, string txtId)
        {
            // ✅ Check if CFL already exists
            bool cflExists = oForm.ChooseFromLists
                .Cast<SAPbouiCOM.ChooseFromList>()
                .Any(cfl => cfl.UniqueID == id);

            if (!cflExists)
            {
                // Create CFL parameters
                var oCFLParams = (SAPbouiCOM.ChooseFromListCreationParams)
                    Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLParams.MultiSelection = false;
                oCFLParams.ObjectType = "2"; // Business Partner
                oCFLParams.UniqueID = id;

                // Add CFL to form
                var oCFL = oForm.ChooseFromLists.Add(oCFLParams);

                // Add conditions
                var conditions = oCFL.GetConditions();
                var condition = conditions.Add();
                condition.Alias = "CardType";
                condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                condition.CondVal = "C"; // Customer
                oCFL.SetConditions(conditions);
            }

            // ✅ Check if DataSource already exists
            if (!oForm.DataSources.UserDataSources.Cast<SAPbouiCOM.UserDataSource>().Any(uds => uds.UID == ds))
            {
                oForm.DataSources.UserDataSources.Add(ds, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            }

            var oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(txtId).Specific;
            oEdit.DataBind.SetBound(true, "", ds);
            oEdit.ChooseFromListUID = id;      // ✅ use your parameter `id` instead of hardcoded "2"
            oEdit.ChooseFromListAlias = "CardCode";
            //// ✅ Check if the item exists before binding
            //if (ItemIsExists(oForm, txtId))
            //{
            //    var oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(txtId).Specific;
            //    oEdit.DataBind.SetBound(true, "", ds);
            //    oEdit.ChooseFromListUID = id;      // ✅ use your parameter `id` instead of hardcoded "2"
            //    oEdit.ChooseFromListAlias = "CardCode";
            //}
        }

        private void SetDocumentCfl(SAPbouiCOM.Form oForm, string id, string ds, string txtId)
        {
            try
            {
                // ✅ Check if CFL already exists
                bool cflExists = oForm.ChooseFromLists
                    .Cast<SAPbouiCOM.ChooseFromList>()
                    .Any(cfl => cfl.UniqueID == id);
                string selectedDoc = SelectedCkBox.Where((d) => d.Value).First().Key;
                if (!cflExists)
                {
                    // Create CFL parameters
                    var oCFLParams = (SAPbouiCOM.ChooseFromListCreationParams)
                        Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                    oCFLParams.MultiSelection = false;
                    switch (selectedDoc)
                    {
                        case "OINV":
                            oCFLParams.ObjectType = "13";
                            break;
                        case "ODPI":
                            oCFLParams.ObjectType = "203";
                            break;
                        case "ORIN":
                            oCFLParams.ObjectType = "14";
                            break;
                        default:
                            break;
                    }
                    oCFLParams.UniqueID = id;

                    // Add CFL to form
                    var oCFL = oForm.ChooseFromLists.Add(oCFLParams);

                    // Add conditions
                    var conditions = oCFL.GetConditions();
                    var condition = conditions.Add();
                    condition.Alias = "CANCELED";
                    condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    condition.CondVal = "N"; // Customer
                    oCFL.SetConditions(conditions);
                }

                // ✅ Check if DataSource already exists
                if (!oForm.DataSources.UserDataSources.Cast<SAPbouiCOM.UserDataSource>().Any(uds => uds.UID == ds))
                {
                    oForm.DataSources.UserDataSources.Add(ds, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
                }

                var oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(txtId).Specific;
                oEdit.DataBind.SetBound(true, "", ds);
                oEdit.ChooseFromListUID = id;    

                //// ✅ Check if the item exists before binding
                //if (ItemIsExists(oForm, txtId))
                //{
                //    var oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(txtId).Specific;
                //    oEdit.DataBind.SetBound(true, "", ds);
                //    oEdit.ChooseFromListUID = id;      // ✅ use your parameter `id` instead of hardcoded "2"
                //                                       //oEdit.ChooseFromListAlias = "CardCode";
                //}
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void SetBranchCfl(SAPbouiCOM.Form oForm, string id, string ds, string txtId)
        {
            // ✅ Check if CFL already exists
            bool cflExists = oForm.ChooseFromLists
                .Cast<SAPbouiCOM.ChooseFromList>()
                .Any(cfl => cfl.UniqueID == id);

            SAPbouiCOM.ChooseFromList oCFL = null;

            if (!cflExists)
            {
                // Create CFL parameters
                var oCFLParams = (SAPbouiCOM.ChooseFromListCreationParams)
                    Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLParams.MultiSelection = false;
                oCFLParams.ObjectType = "119";
                oCFLParams.UniqueID = id;

                // Add CFL to form
                oCFL = oForm.ChooseFromLists.Add(oCFLParams);

                // Optional: add conditions here if needed
            }
            else
            {
                oCFL = oForm.ChooseFromLists.Item(id);
            }

            // ✅ Check if DataSource already exists
            if (!oForm.DataSources.UserDataSources.Cast<SAPbouiCOM.UserDataSource>().Any(uds => uds.UID == ds))
            {
                oForm.DataSources.UserDataSources.Add(ds, SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 20);
            }

            var oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(txtId).Specific;
            oEdit.DataBind.SetBound(true, "", ds);
            oEdit.ChooseFromListUID = id;

            //// ✅ Bind EditText to DataSource and CFL
            //if (ItemIsExists(oForm, txtId))
            //{
            //    var oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(txtId).Specific;
            //    oEdit.DataBind.SetBound(true, "", ds);
            //    oEdit.ChooseFromListUID = id;
            //    //oEdit.ChooseFromListAlias = "Code"; // or "BPLName" if you want the name
            //}
        }

        private bool ItemIsExists(SAPbouiCOM.Form oForm, string itemUid)
        {
            try
            {
                for (int i = 0; i < oForm.Items.Count; i++)
                {
                    var c = oForm.Items.Item(i).UniqueID;
                    if (oForm.Items.Item(i).UniqueID == itemUid) // 1-based index
                        return true;
                }
            }
            catch (Exception)
            {

                throw;
            }

            return false;
        }

        private bool DsIsExists(SAPbouiCOM.Form oForm, string dsUid)
        {
            try
            {
                for (int i = 0; i < oForm.DataSources.UserDataSources.Count; i++)
                {
                    var c = oForm.DataSources.UserDataSources.Item(i).UID;
                    if (oForm.DataSources.UserDataSources.Item(i).UID == dsUid) // 1-based index
                        return true;
                }
            }
            catch (Exception)
            {

                throw;
            }

            return false;
        }

        private bool DBSIsExists(SAPbouiCOM.Form oForm, string tableName)
        {
            try
            {
                for (int i = 0; i < oForm.DataSources.DBDataSources.Count; i++)
                {
                    var c = oForm.DataSources.DBDataSources.Item(i).TableName;
                    if (oForm.DataSources.DBDataSources.Item(i).TableName == tableName) // 1-based index
                        return true;
                }
            }
            catch (Exception)
            {

                throw;
            }

            return false;
        }

        private bool DtIsExists(SAPbouiCOM.Form oForm, string dsUid)
        {
            try
            {
                for (int i = 0; i < oForm.DataSources.DataTables.Count; i++)
                {
                    var c = oForm.DataSources.DataTables.Item(i).UniqueID;
                    if (oForm.DataSources.DataTables.Item(i).UniqueID == dsUid) // 1-based index
                        return true;
                }
            }
            catch (Exception)
            {

                throw;
            }

            return false;
        }

        private void MapModelToUI(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);

                if (_pb != null) { _pb.Stop(); _pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Loading form data...", 0, false);

                if (coretaxModel == null)
                    return;

                // ==== Header ====
                SetValueDS(oForm, "CkInvDS", coretaxModel.IsARInvoice ? "Y":"N");
                SelectedCkBox["OINV"] = coretaxModel.IsARInvoice;
                SetValueDS(oForm, "CkDpDS", coretaxModel.IsARDownPayment ? "Y" : "N");
                SelectedCkBox["ODPI"] = coretaxModel.IsARDownPayment;
                SetValueDS(oForm, "CkCmDS", coretaxModel.IsARCreditMemo ? "Y" : "N");
                SelectedCkBox["ORIN"] = coretaxModel.IsARCreditMemo;

                if (coretaxModel.FromDate.HasValue) SetValueEdit(oForm, "TFromDt", coretaxModel.FromDate.Value.ToString("yyyyMMdd"));
                if (coretaxModel.ToDate.HasValue) SetValueEdit(oForm, "TToDt", coretaxModel.ToDate.Value.ToString("yyyyMMdd"));
                if (coretaxModel.FromDocNum > 0) SetValueEdit(oForm, "TFromDoc", coretaxModel.FromDocNum.ToString());
                if (coretaxModel.ToDocNum > 0) SetValueEdit(oForm, "TToDoc", coretaxModel.ToDocNum.ToString());
                if (!string.IsNullOrEmpty(coretaxModel.FromCust)) SetValueEdit(oForm, "TCustFrom", coretaxModel.FromCust);
                if (!string.IsNullOrEmpty(coretaxModel.ToCust)) SetValueEdit(oForm, "TCustTo", coretaxModel.ToCust);
                if (!string.IsNullOrEmpty(coretaxModel.FromBranch)) SetValueCb(oForm, "CbBrFrom", coretaxModel.FromBranch);
                if (!string.IsNullOrEmpty(coretaxModel.ToBranch)) SetValueEdit(oForm, "CbBrTo", coretaxModel.ToBranch);
                if (!string.IsNullOrEmpty(coretaxModel.FromOutlet)) SetValueCb(oForm, "CbOtlFrom", coretaxModel.FromOutlet);
                if (!string.IsNullOrEmpty(coretaxModel.ToOutlet)) SetValueEdit(oForm, "CbOtlTo", coretaxModel.ToOutlet);

                if (coretaxModel.DocNum != null) SetValueEdit(oForm, "TDocNum", coretaxModel.DocNum.ToString());
                if (!string.IsNullOrEmpty(coretaxModel.Status))
                    SetValueEdit(oForm, "TStatus", coretaxModel.Status == "O" ? "Open" : "Closed");
                if (!string.IsNullOrEmpty(coretaxModel.SeriesName)) SetValueEdit(oForm, "TSeries", coretaxModel.SeriesName);
                if (coretaxModel.DocDate.HasValue) SetValueEdit(oForm, "TDocDt", coretaxModel.DocDate.Value.ToString("yyyyMMdd"));
                if (coretaxModel.PostDate.HasValue) SetValueEdit(oForm, "TPostDt", coretaxModel.PostDate.Value.ToString("yyyyMMdd"));
                
                // ==== Details ====
                if (coretaxModel.Detail != null && coretaxModel.Detail.Any())
                {
                    var gInvList = coretaxModel.Detail
                        .GroupBy(p => new
                        {
                            p.DocEntry,
                            p.NoDocument,
                            p.BPCode,
                            p.BPName,
                            p.ObjectType,
                            p.ObjectName,
                            p.InvDate,
                            p.BranchCode,
                            p.BranchName,
                            p.OutletCode,
                            p.OutletName
                        })
                        .Select(g => new FilterDataModel
                        {
                            DocEntry = g.Key.DocEntry,
                            DocNo = g.Key.NoDocument,
                            CardCode = g.Key.BPCode,
                            CardName = g.Key.BPName,
                            ObjType = g.Key.ObjectType,
                            ObjName = g.Key.ObjectName,
                            PostDate = g.Key.InvDate,
                            BranchCode = g.Key.BranchCode,
                            BranchName = g.Key.BranchName,
                            OutletCode = g.Key.OutletCode,
                            OutletName = g.Key.OutletName,
                            Selected = true // default checked
                    })
                        .ToList();

                    if (gInvList.Any())
                    {
                        FindListModel = gInvList;
                        SetMtFind(oForm);     // fast load matrix
                        if (coretaxModel.Status == "O" && coretaxModel.DocEntry != null && coretaxModel.DocEntry != 0)
                        {
                            oForm.Items.Item("BtCSV").Enabled = true;
                            oForm.Items.Item("BtXML").Enabled = true;
                        }
                        else
                        {
                            oForm.Items.Item("BtCSV").Enabled = false;
                            oForm.Items.Item("BtXML").Enabled = false;
                        }
                    }
                    else
                    {
                        oForm.Items.Item("BtCSV").Enabled = false;
                        oForm.Items.Item("BtXML").Enabled = false;
                    }
                }
                else
                {
                    ClearMatrix(oForm, "MtFind", "DT_FILTER");
                }
                SetMtGenerate(oForm);
                // Change button caption to Update after first save
                var btnUpdate = (SAPbouiCOM.Button)oForm.Items.Item("BtSave").Specific;
                if (coretaxModel.Status == "O" && coretaxModel.DocEntry != null && coretaxModel.DocEntry != 0)
                {
                    btnUpdate.Caption = "Update";
                    //oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                }
                else
                {
                    btnUpdate.Caption = "Add";
                    //oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                }
                UpdateFormItemStates(oForm, coretaxModel, FindListModel);
                SetBtnExport(oForm);
            }
            catch (Exception e)
            {
                Application.SBO_Application.StatusBar.SetText($"Error in MapModelToUI: {e.Message}", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                throw;
            }
            finally
            {
                oForm.Freeze(false);
                if (_pb != null) { _pb.Stop(); _pb = null; }
            }
        }

        private void UpdateFormItemStates(SAPbouiCOM.Form oForm, CoretaxModel coretaxModel, List<FilterDataModel> FindListModel)
        {
            // Common item groups
            string[] docItems = { "TFromDoc", "TToDoc", "CkAllDoc" };
            string[] dtItems = { "TFromDt", "TToDt", "CkAllDt" };
            string[] custItems = { "TCustFrom", "TCustTo", "CkAllCust" };
            string[] brItems = { "CbBrFrom", "CbBrTo", "CkAllBr" };
            string[] otlItems = { "CbOtlFrom", "CbOtlTo", "CkAllOtl" };
            string[] actionButtons = { "BtFind", "BtGen", "BtSave", "BtClose" };
            string[] chkItems = { "CkInv", "CkDp", "CkCM" };

            bool isOpen = (coretaxModel.Status == "O");
            bool hasDocEntry = (coretaxModel.DocEntry ?? 0) != 0;

            if (isOpen)
            {
                // Always enable invoice type checkboxes
                SetEnabled(oForm, chkItems, true);

                if (hasDocEntry)
                {
                    int selectedCount = new[]
                    {
                coretaxModel.IsARInvoice,
                coretaxModel.IsARDownPayment,
                coretaxModel.IsARCreditMemo
            }.Count(b => b);

                    // Doc filters only enabled if exactly 1 type selected
                    SetEnabled(oForm, docItems, selectedCount == 1);

                    // Always enable other filters + action buttons
                    SetEnabled(oForm, dtItems, true);
                    SetEnabled(oForm, custItems, true);
                    SetEnabled(oForm, brItems, true);
                    SetEnabled(oForm, otlItems, true);

                    SetEnabled(oForm, actionButtons, true);

                    // BtGen depends on selected FindList
                    oForm.Items.Item("BtGen").Enabled = (FindListModel != null && FindListModel.Any(f => f.Selected));
                }
                else
                {
                    // No DocEntry yet → disable all except checkboxes
                    SetEnabled(oForm, docItems, false);
                    SetEnabled(oForm, dtItems, false);
                    SetEnabled(oForm, custItems, false);
                    SetEnabled(oForm, brItems, false);
                    SetEnabled(oForm, otlItems, false);
                    SetEnabled(oForm, actionButtons, false);
                }
            }
            else
            {
                // Status != "O" → disable everything
                SetEnabled(oForm, chkItems, false);
                SetEnabled(oForm, docItems, false);
                SetEnabled(oForm, dtItems, false);
                SetEnabled(oForm, custItems, false);
                SetEnabled(oForm, brItems, false);
                SetEnabled(oForm, otlItems, false);
                SetEnabled(oForm, actionButtons, false);
                SetEnabled(oForm, chkItems, false);
            }
        }

        private void SetEnabled(SAPbouiCOM.Form oForm, string[] itemIds, bool enabled)
        {
            foreach (var id in itemIds)
            {
                if (!enabled && oForm.ActiveItem == id)
                {
                    RemoveFocus(oForm);
                }

                oForm.Items.Item(id).Enabled = enabled;
            }
        }



        private void SetValueEdit(SAPbouiCOM.Form oForm, string id, string val)
        {
            oForm.Freeze(true);
            try
            {
                //if (!ItemIsExists(oForm, id)) return;
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(id).Specific;
                oEdit.Value = val;
                oEdit.Active = false;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void SetValueCheck(SAPbouiCOM.Form oForm, string id, bool val)
        {
            oForm.Freeze(true);
            try
            {
                //if (!ItemIsExists(oForm, id)) return;
                SAPbouiCOM.CheckBox oCheck = (SAPbouiCOM.CheckBox)oForm.Items.Item(id).Specific;
                oCheck.Checked = val;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void SetValueCb(SAPbouiCOM.Form oForm, string id, string val)
        {
            oForm.Freeze(true);
            try
            {
                //if (!ItemIsExists(oForm, id)) return;
                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item(id).Specific;
                oCombo.Select(val, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oCombo.Active = false;
            }
            finally
            {
                oForm.Freeze(false);
            }
        }

        private void SetValueDS(SAPbouiCOM.Form oForm, string id, string val)
        {
            if (!DsIsExists(oForm, id)) return;
            oForm.DataSources.UserDataSources.Item(id).Value = val;
        }

        private void ResetHeader(SAPbouiCOM.Form oForm)
        {
            coretaxModel = new CoretaxModel();

            SetValueDS(oForm, "CkInvDS", "N");
            SetValueDS(oForm, "CkDpDS", "N");
            SetValueDS(oForm, "CkCmDS", "N");

            SetValueDS(oForm, "DtFromDS", null);
            SetValueDS(oForm, "DtToDS", null);
            SetValueDS(oForm, "DocDtDS", null);
            SetValueDS(oForm, "PostDtDS", null);

            ClearEdit(oForm, "TFromDoc");
            ClearEdit(oForm, "TToDoc");
            ClearEdit(oForm, "TCustFrom");
            ClearEdit(oForm, "TCustTo");
            ClearEdit(oForm, "TSeries");
            ClearEdit(oForm, "TDocNum");
            ClearEdit(oForm, "TStatus");

            ClearCombo(oForm, "CbBrFrom");
            ClearCombo(oForm, "CbOtlFrom");
            ClearCombo(oForm, "CbOtlTo");
        }

        private void Reset(SAPbouiCOM.Form oForm)
        {
            try
            {
                ResetHeader(oForm);

                //ClearCheckBox(oForm,"CkInv", "CkInvDS");
                //ClearCheckBox(oForm,"CkDp", "CkDpDS");
                //ClearCheckBox(oForm,"CkCM", "CkCmDS");

                //ClearEditDate(oForm, "TFromDt", "DtFromDS");
                //ClearEditDate(oForm, "TToDt", "DtToDS");
                //ClearEditDate(oForm, "TDocDt", "DocDtDS");
                //ClearEditDate(oForm, "TPostDt", "PostDtDS");

                //ClearEdit(oForm, "TFromDoc");
                //ClearEdit(oForm, "TToDoc");
                //ClearEdit(oForm, "TCustFrom");
                //ClearEdit(oForm, "TCustTo");
                //ClearEdit(oForm, "TSeries");
                //ClearEdit(oForm, "TDocNum");
                //ClearEdit(oForm, "TStatus");

                //ClearCombo(oForm, "CbBrFrom");
                //ClearCombo(oForm, "CbOtlFrom");
                //ClearCombo(oForm, "CbOtlTo");

                ResetDetail(oForm);
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void ResetDetail(SAPbouiCOM.Form oForm)
        {
            if(FindListModel != null && FindListModel.Any())
            {
                FindListModel.Clear();
                ClearMatrix(oForm, "MtFind", "DT_FILTER");
            }
            if (coretaxModel.Detail != null && coretaxModel.Detail.Any())
            {
                coretaxModel.Detail.Clear();
                ClearMatrix(oForm, "MtDetail", "DT_DETAIL", "@T2_CORETAX");
            }
        }
        
        private void SetMtGenerate(SAPbouiCOM.Form oForm)
        {
            if ((coretaxModel.DocEntry ?? 0) == 0) return;
            int docEntry = coretaxModel.DocEntry ?? 0;
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MtDetail").Specific;

            if (!DBSIsExists(oForm, "@T2_CORETAX_DT"))
            {
                oForm.DataSources.DBDataSources.Add("@T2_CORETAX_DT");
            }

            ClearMatrix(oForm, "MtDetail", "DT_DETAIL", "@T2_CORETAX_DT");

            SAPbouiCOM.DBDataSource oDBDS = oForm.DataSources.DBDataSources.Item("@T2_CORETAX_DT");
            
            oMatrix.Columns.Item("Col_1").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Inv_Date");
            oMatrix.Columns.Item("Col_1").Width = 80;
            oMatrix.Columns.Item("Col_2").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_No_Doc");
            oMatrix.Columns.Item("Col_2").Width = 80;
            oMatrix.Columns.Item("Col_3").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Object_Type");
            oMatrix.Columns.Item("Col_3").Width = 80;
            oMatrix.Columns.Item("Col_4").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Object_Name");
            oMatrix.Columns.Item("Col_4").Width = 80;
            oMatrix.Columns.Item("Col_5").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_BP_Code");
            oMatrix.Columns.Item("Col_5").Width = 80;
            oMatrix.Columns.Item("Col_6").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_BP_Name");
            oMatrix.Columns.Item("Col_6").Width = 100;
            oMatrix.Columns.Item("Col_7").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Seller_IDTKU");
            oMatrix.Columns.Item("Col_7").Width = 100;
            oMatrix.Columns.Item("Col_8").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Buyer_Doc");
            oMatrix.Columns.Item("Col_8").Width = 80;
            oMatrix.Columns.Item("Col_9").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Nomor_NPWP");
            oMatrix.Columns.Item("Col_9").Width = 80;
            oMatrix.Columns.Item("Col_10").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_NPWP_Name");
            oMatrix.Columns.Item("Col_10").Width = 100;
            oMatrix.Columns.Item("Col_11").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_NPWP_Address");
            oMatrix.Columns.Item("Col_11").Width = 120;
            oMatrix.Columns.Item("Col_12").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Buyer_IDTKU");
            oMatrix.Columns.Item("Col_12").Width = 80;
            oMatrix.Columns.Item("Col_13").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Item_Code");
            oMatrix.Columns.Item("Col_13").Width = 80;
            oMatrix.Columns.Item("Col_14").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Item_Name");
            oMatrix.Columns.Item("Col_14").Width = 100;
            oMatrix.Columns.Item("Col_15").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Item_Unit");
            oMatrix.Columns.Item("Col_15").Width = 80;
            oMatrix.Columns.Item("Col_16").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Item_Price");
            oMatrix.Columns.Item("Col_16").Width = 80;
            oMatrix.Columns.Item("Col_17").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Qty");
            oMatrix.Columns.Item("Col_17").Width = 80;
            oMatrix.Columns.Item("Col_18").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Total_Disc");
            oMatrix.Columns.Item("Col_18").Width = 80;
            oMatrix.Columns.Item("Col_19").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Tax_Base");
            oMatrix.Columns.Item("Col_19").Width = 80;
            oMatrix.Columns.Item("Col_20").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Other_Tax_Base");
            oMatrix.Columns.Item("Col_20").Width = 80;
            oMatrix.Columns.Item("Col_21").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_VAT_Rate");
            oMatrix.Columns.Item("Col_21").Width = 80;
            oMatrix.Columns.Item("Col_22").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Amount_VAT");
            oMatrix.Columns.Item("Col_22").Width = 80;
            oMatrix.Columns.Item("Col_23").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_STLG_Rate");
            oMatrix.Columns.Item("Col_23").Width = 80;
            oMatrix.Columns.Item("Col_24").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_STLG");
            oMatrix.Columns.Item("Col_24").Width = 80;
            oMatrix.Columns.Item("Col_25").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Jenis_Pajak");
            oMatrix.Columns.Item("Col_25").Width = 80;
            oMatrix.Columns.Item("Col_26").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Add_Info");
            oMatrix.Columns.Item("Col_26").Width = 100;
            oMatrix.Columns.Item("Col_27").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Ket_Tambahan");
            oMatrix.Columns.Item("Col_27").Width = 100;
            oMatrix.Columns.Item("Col_28").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Pajak_Pengganti");
            oMatrix.Columns.Item("Col_28").Width = 80;
            oMatrix.Columns.Item("Col_29").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Referensi");
            oMatrix.Columns.Item("Col_29").Width = 80;
            oMatrix.Columns.Item("Col_30").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Status");
            oMatrix.Columns.Item("Col_30").Width = 80;
            oMatrix.Columns.Item("Col_31").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Kode_Dok_Pendukung");
            oMatrix.Columns.Item("Col_31").Width = 80;
            oMatrix.Columns.Item("Col_32").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Buyer_Country");
            oMatrix.Columns.Item("Col_32").Width = 80;
            oMatrix.Columns.Item("Col_33").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Buyer_Email");
            oMatrix.Columns.Item("Col_33").Width = 80;
            oMatrix.Columns.Item("Col_34").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Branch_Code");
            oMatrix.Columns.Item("Col_34").Width = 80;
            oMatrix.Columns.Item("Col_35").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Branch_Name");
            oMatrix.Columns.Item("Col_35").Width = 80;
            oMatrix.Columns.Item("Col_36").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Outlet_Code");
            oMatrix.Columns.Item("Col_36").Width = 80;
            oMatrix.Columns.Item("Col_37").DataBind.SetBound(true, "@T2_CORETAX_DT", "U_T2_Outlet_Name");
            oMatrix.Columns.Item("Col_37").Width = 80;
            oMatrix.Columns.Item("#").DataBind.SetBound(true, "@T2_CORETAX_DT", "LineId");
            oMatrix.Columns.Item("#").Width = 30;

            // buat objek conditions kosong
            SAPbouiCOM.Conditions oConds = (SAPbouiCOM.Conditions)Application.SBO_Application.CreateObject(
                SAPbouiCOM.BoCreatableObjectType.cot_Conditions
            );

            // tambahkan 1 condition
            SAPbouiCOM.Condition oCond = oConds.Add();
            oCond.Alias = "DocEntry"; // nama field di UDT
            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
            oCond.CondVal = docEntry.ToString();
            oDBDS.Query(oConds);

            // tampilkan hasil ke matrix
            oMatrix.Clear();
            
            oMatrix.LoadFromDataSource();

            int white = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("#").Cells.Item(i).Specific).Value = i.ToString();
                oMatrix.CommonSetting.SetRowBackColor(i, white);
            }

            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            oMatrix.AutoResizeColumns();
        }

        private void SetMtGenerateByModel(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MtDetail").Specific;
                // Create DataTable if not exists
                SAPbouiCOM.DataTable oDT;
                if (!DtIsExists(oForm, "DT_DETAIL"))
                {
                    oDT = oForm.DataSources.DataTables.Add("DT_DETAIL");
                }
                else
                {
                    oDT = oForm.DataSources.DataTables.Item("DT_DETAIL");
                }
                // Clear previous rows
                oDT.Clear();

                // Also clear matrix (important)
                oMatrix.Clear();

                // Define all columns (make sure sizes are large enough)
                oDT.Columns.Add("InvDate", SAPbouiCOM.BoFieldsType.ft_Date);
                oDT.Columns.Add("NoDoc", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("ObjType", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("ObjName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                oDT.Columns.Add("BPCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("BPName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                oDT.Columns.Add("SellerIDTKU", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("BuyerDoc", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("NomorNPWP", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("NPWPName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                oDT.Columns.Add("NPWPAddress", SAPbouiCOM.BoFieldsType.ft_Text, 150);
                oDT.Columns.Add("BuyerIDTKU", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("ItemName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                oDT.Columns.Add("ItemUnit", SAPbouiCOM.BoFieldsType.ft_Text, 20);
                oDT.Columns.Add("ItemPrice", SAPbouiCOM.BoFieldsType.ft_Price);
                oDT.Columns.Add("Qty", SAPbouiCOM.BoFieldsType.ft_Quantity);
                oDT.Columns.Add("TotalDisc", SAPbouiCOM.BoFieldsType.ft_Price);
                oDT.Columns.Add("TaxBase", SAPbouiCOM.BoFieldsType.ft_Price);
                oDT.Columns.Add("OtherTaxBase", SAPbouiCOM.BoFieldsType.ft_Price);
                oDT.Columns.Add("VATRate", SAPbouiCOM.BoFieldsType.ft_Percent);
                oDT.Columns.Add("AmountVAT", SAPbouiCOM.BoFieldsType.ft_Price);
                oDT.Columns.Add("STLGRate", SAPbouiCOM.BoFieldsType.ft_Percent);
                oDT.Columns.Add("STLG", SAPbouiCOM.BoFieldsType.ft_Price);
                oDT.Columns.Add("JenisPajak", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("AddInfo", SAPbouiCOM.BoFieldsType.ft_Text, 150);
                oDT.Columns.Add("KetTambahan", SAPbouiCOM.BoFieldsType.ft_Text, 150);
                oDT.Columns.Add("PajakPengganti", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("Referensi", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("Status", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("KodeDokPendukung", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("BuyerCountry", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("BuyerEmail", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                oDT.Columns.Add("BranchCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("BranchName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                oDT.Columns.Add("OutletCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("OutletName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                oDT.Columns.Add("#", SAPbouiCOM.BoFieldsType.ft_Integer);


                // Fill DataTable from model
                for (int i = 0; i < coretaxModel.Detail.Count; i++)
                {
                    var row = coretaxModel.Detail[i];
                    oDT.Rows.Add();

                    // Convert date string to SAP date format
                    if (!string.IsNullOrEmpty(row.InvDate) && DateTime.TryParse(row.InvDate, out var invDate))
                        oDT.SetValue("InvDate", i, invDate.ToString("yyyyMMdd"));
                    else
                        oDT.SetValue("InvDate", i, "");

                    oDT.SetValue("NoDoc", i, row.NoDocument ?? "");
                    oDT.SetValue("ObjType", i, row.ObjectType ?? "");
                    oDT.SetValue("ObjName", i, row.ObjectName ?? "");
                    oDT.SetValue("BPCode", i, row.BPCode ?? "");
                    oDT.SetValue("BPName", i, row.BPName ?? "");
                    oDT.SetValue("SellerIDTKU", i, row.SellerIDTKU ?? "");
                    oDT.SetValue("BuyerDoc", i, row.BuyerDocument ?? "");
                    oDT.SetValue("NomorNPWP", i, row.NomorNPWP ?? "");
                    oDT.SetValue("NPWPName", i, row.NPWPName ?? "");
                    oDT.SetValue("NPWPAddress", i, row.NPWPAddress ?? "");
                    oDT.SetValue("BuyerIDTKU", i, row.BuyerIDTKU ?? "");
                    oDT.SetValue("ItemCode", i, row.ItemCode ?? "");
                    oDT.SetValue("ItemName", i, row.ItemName ?? "");
                    oDT.SetValue("ItemUnit", i, row.ItemUnit ?? "");
                    oDT.SetValue("ItemPrice", i, row.ItemPrice);
                    oDT.SetValue("Qty", i, row.Qty);
                    oDT.SetValue("TotalDisc", i, row.TotalDisc);
                    oDT.SetValue("TaxBase", i, row.TaxBase);
                    oDT.SetValue("OtherTaxBase", i, row.OtherTaxBase);
                    oDT.SetValue("VATRate", i, row.VATRate);
                    oDT.SetValue("AmountVAT", i, row.AmountVAT);
                    oDT.SetValue("STLGRate", i, row.STLGRate);
                    oDT.SetValue("STLG", i, row.STLG);
                    oDT.SetValue("JenisPajak", i, row.JenisPajak ?? "");
                    oDT.SetValue("AddInfo", i, row.AddInfo ?? "");
                    oDT.SetValue("KetTambahan", i, row.KetTambahan ?? "");
                    oDT.SetValue("PajakPengganti", i, row.PajakPengganti ?? "");
                    oDT.SetValue("Referensi", i, row.Referensi ?? "");
                    oDT.SetValue("Status", i, row.Status ?? "");
                    oDT.SetValue("KodeDokPendukung", i, row.KodeDokumenPendukung ?? "");
                    oDT.SetValue("BuyerCountry", i, row.BuyerCountry ?? "");
                    oDT.SetValue("BuyerEmail", i, row.BuyerEmail ?? "");
                    oDT.SetValue("BranchCode", i, row.BranchCode ?? "");
                    oDT.SetValue("BranchName", i, row.BranchName ?? "");
                    oDT.SetValue("OutletCode", i, row.OutletCode ?? "");
                    oDT.SetValue("OutletName", i, row.OutletName ?? "");
                    oDT.SetValue("#", i, (i + 1));
                }
                
                oMatrix.Columns.Item("Col_1").DataBind.Bind("DT_DETAIL", "InvDate");
                oMatrix.Columns.Item("Col_1").Width = 80;
                oMatrix.Columns.Item("Col_2").DataBind.Bind("DT_DETAIL", "NoDoc");
                oMatrix.Columns.Item("Col_2").Width = 80;
                oMatrix.Columns.Item("Col_3").DataBind.Bind("DT_DETAIL", "ObjType");
                oMatrix.Columns.Item("Col_3").Width = 80;
                oMatrix.Columns.Item("Col_4").DataBind.Bind("DT_DETAIL", "ObjName");
                oMatrix.Columns.Item("Col_4").Width = 100;
                oMatrix.Columns.Item("Col_5").DataBind.Bind("DT_DETAIL", "BPCode");
                oMatrix.Columns.Item("Col_5").Width = 80;
                oMatrix.Columns.Item("Col_6").DataBind.Bind("DT_DETAIL", "BPName");
                oMatrix.Columns.Item("Col_6").Width = 100;
                oMatrix.Columns.Item("Col_7").DataBind.Bind("DT_DETAIL", "SellerIDTKU");
                oMatrix.Columns.Item("Col_7").Width = 80;
                oMatrix.Columns.Item("Col_8").DataBind.Bind("DT_DETAIL", "BuyerDoc");
                oMatrix.Columns.Item("Col_8").Width = 80;
                oMatrix.Columns.Item("Col_9").DataBind.Bind("DT_DETAIL", "NomorNPWP");
                oMatrix.Columns.Item("Col_9").Width = 80;
                oMatrix.Columns.Item("Col_10").DataBind.Bind("DT_DETAIL", "NPWPName");
                oMatrix.Columns.Item("Col_10").Width = 80;
                oMatrix.Columns.Item("Col_11").DataBind.Bind("DT_DETAIL", "NPWPAddress");
                oMatrix.Columns.Item("Col_11").Width = 120;
                oMatrix.Columns.Item("Col_12").DataBind.Bind("DT_DETAIL", "BuyerIDTKU");
                oMatrix.Columns.Item("Col_12").Width = 80;
                oMatrix.Columns.Item("Col_13").DataBind.Bind("DT_DETAIL", "ItemCode");
                oMatrix.Columns.Item("Col_13").Width = 80;
                oMatrix.Columns.Item("Col_14").DataBind.Bind("DT_DETAIL", "ItemName");
                oMatrix.Columns.Item("Col_14").Width = 100;
                oMatrix.Columns.Item("Col_15").DataBind.Bind("DT_DETAIL", "ItemUnit");
                oMatrix.Columns.Item("Col_15").Width = 80;
                oMatrix.Columns.Item("Col_16").DataBind.Bind("DT_DETAIL", "ItemPrice");
                oMatrix.Columns.Item("Col_16").Width = 80;
                oMatrix.Columns.Item("Col_17").DataBind.Bind("DT_DETAIL", "Qty");
                oMatrix.Columns.Item("Col_17").Width = 80;
                oMatrix.Columns.Item("Col_18").DataBind.Bind("DT_DETAIL", "TotalDisc");
                oMatrix.Columns.Item("Col_18").Width = 80;
                oMatrix.Columns.Item("Col_19").DataBind.Bind("DT_DETAIL", "TaxBase");
                oMatrix.Columns.Item("Col_19").Width = 80;
                oMatrix.Columns.Item("Col_20").DataBind.Bind("DT_DETAIL", "OtherTaxBase");
                oMatrix.Columns.Item("Col_20").Width = 80;
                oMatrix.Columns.Item("Col_21").DataBind.Bind("DT_DETAIL", "VATRate");
                oMatrix.Columns.Item("Col_21").Width = 80;
                oMatrix.Columns.Item("Col_22").DataBind.Bind("DT_DETAIL", "AmountVAT");
                oMatrix.Columns.Item("Col_22").Width = 80;
                oMatrix.Columns.Item("Col_23").DataBind.Bind("DT_DETAIL", "STLGRate");
                oMatrix.Columns.Item("Col_23").Width = 80;
                oMatrix.Columns.Item("Col_24").DataBind.Bind("DT_DETAIL", "STLG");
                oMatrix.Columns.Item("Col_24").Width = 80;
                oMatrix.Columns.Item("Col_25").DataBind.Bind("DT_DETAIL", "JenisPajak");
                oMatrix.Columns.Item("Col_25").Width = 80;
                oMatrix.Columns.Item("Col_26").DataBind.Bind("DT_DETAIL", "AddInfo");
                oMatrix.Columns.Item("Col_26").Width = 80;
                oMatrix.Columns.Item("Col_27").DataBind.Bind("DT_DETAIL", "KetTambahan");
                oMatrix.Columns.Item("Col_27").Width = 80;
                oMatrix.Columns.Item("Col_28").DataBind.Bind("DT_DETAIL", "PajakPengganti");
                oMatrix.Columns.Item("Col_28").Width = 80;
                oMatrix.Columns.Item("Col_29").DataBind.Bind("DT_DETAIL", "Referensi");
                oMatrix.Columns.Item("Col_29").Width = 80;
                oMatrix.Columns.Item("Col_30").DataBind.Bind("DT_DETAIL", "Status");
                oMatrix.Columns.Item("Col_30").Width = 80;
                oMatrix.Columns.Item("Col_31").DataBind.Bind("DT_DETAIL", "KodeDokPendukung");
                oMatrix.Columns.Item("Col_31").Width = 80;
                oMatrix.Columns.Item("Col_32").DataBind.Bind("DT_DETAIL", "BuyerCountry");
                oMatrix.Columns.Item("Col_32").Width = 80;
                oMatrix.Columns.Item("Col_33").DataBind.Bind("DT_DETAIL", "BuyerEmail");
                oMatrix.Columns.Item("Col_33").Width = 80;
                oMatrix.Columns.Item("Col_34").DataBind.Bind("DT_DETAIL", "BranchCode");
                oMatrix.Columns.Item("Col_34").Width = 80;
                oMatrix.Columns.Item("Col_35").DataBind.Bind("DT_DETAIL", "BranchName");
                oMatrix.Columns.Item("Col_35").Width = 100;
                oMatrix.Columns.Item("Col_36").DataBind.Bind("DT_DETAIL", "OutletCode");
                oMatrix.Columns.Item("Col_36").Width = 80;
                oMatrix.Columns.Item("Col_37").DataBind.Bind("DT_DETAIL", "OutletName");
                oMatrix.Columns.Item("Col_37").Width = 100;
                oMatrix.Columns.Item("#").DataBind.Bind("DT_DETAIL", "#");
                oMatrix.Columns.Item("#").Width = 30;

                // Load data into matrix
                oMatrix.LoadFromDataSource();
                oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                int white = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                for (int i = 0; i < oMatrix.RowCount; i++)
                {
                    oMatrix.CommonSetting.SetRowBackColor(i + 1, white);
                }
                oMatrix.AutoResizeColumns();

            }
            catch (Exception)
            {

                throw;
            }finally
            {
                oForm.Freeze(false);
            }
        }

    }
}
