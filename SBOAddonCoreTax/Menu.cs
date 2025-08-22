using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
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
        private SAPbouiCOM.DataTable DtGridFind;
        private SAPbouiCOM.DataTable DtGridRes;
        private List<FilterDataModel> FindListModel = new List<FilterDataModel>();
        private TaxInvoiceBulk taxInvoiceBulk = new TaxInvoiceBulk();
        private string FORMUID = "";
        
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
            catch (Exception e)
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
            catch (Exception er)
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
            if (EventInfo.FormUID == FORMUID)
            {
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FORMUID);
                ShowPopupList(oForm);
                Application.SBO_Application.SetStatusBarMessage("Right click event",SAPbouiCOM.BoMessageTime.bmt_Short,false);
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
                            
                            if (DateTime.TryParseExact(oEdit.Value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                                coretaxModel.FromDate = parsedDate;
                        }
                        if (pVal.ItemUID == "TToDt")
                        {
                            SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("TToDt").Specific;

                            if (DateTime.TryParseExact(oEdit.Value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                                coretaxModel.ToDate = parsedDate;
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                        SetBoldUnderlinedLabel(oForm, "lDisp");
                        SetBoldUnderlinedLabel(oForm, "lParam");
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
                        if (pVal.ItemUID == "BtAdd")
                        {
                            var model = TransactionService.GetCoretaxByKey(1);
                            //if (coretaxModel != null)
                            //{
                            //    // Show confirmation dialog
                            //    int response = Application.SBO_Application.MessageBox(
                            //        "Are you sure you want to continue?",  // Message text
                            //        1,                                     // Number of buttons (1=Yes/No, 2=Yes/No/Cancel)
                            //        "Yes",                                 // Button 1 text
                            //        "No",                                  // Button 2 text
                            //        ""                                     // Button 3 text (optional)
                            //    );
                            //    if (response == 1) // Yes
                            //    {
                            //        if (!TransactionService.AddDataCoretax(coretaxModel))
                            //            throw new Exception("Failed to Add Document.");
                            //        Application.SBO_Application.MessageBox("Successfully saved.");
                            //    }
                            //}
                        }

                        //
                        if (pVal.ItemUID == "TFromDt")
                        {
                            coretaxModel.FromDate = EditTextToDate(FormUID, "TFromDt");
                        }
                        if (pVal.ItemUID == "TToDt")
                        {
                            coretaxModel.ToDate = EditTextToDate(FormUID, "TToDt");
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT && !pVal.BeforeAction)
                    {
                        
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && !pVal.BeforeAction)
                    {
                        CflHandler(pVal, FormUID);
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "GdFind" && pVal.ColUID == "Select") // Header click
                        {
                            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                            ClearGrid(oForm, "GdRes", "DT_RES");
                            oForm.Items.Item("BtAdd").Enabled = false;
                            if (pVal.Row == -1)
                            {
                                SelectAllDocHandler(FormUID, "GdFind");
                            }
                            else
                            {
                                SelectSingleDocHandler(FormUID, "GdFind", pVal.Row);
                            }

                            SAPbouiCOM.Item btGen = oForm.Items.Item("BtGen");
                            if (FindListModel != null && FindListModel.Any((f)=>f.Selected))
                            {
                                btGen.Enabled = true;
                            }
                            else
                            {
                                btGen.Enabled = false;
                            }
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
        
        private void AddRightClickMenu(SAPbouiCOM.Form oForm)
        {
            if (!oForm.Menu.Exists("CTX_RIGHT_CLICK"))
            {
                SAPbouiCOM.MenuCreationParams oParams =
                    (SAPbouiCOM.MenuCreationParams)Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);

                oParams.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oParams.UniqueID = "CTX_RIGHT_CLICK";
                oParams.String = "Action List";
                oParams.Position = -1;

                oForm.Menu.AddEx(oParams);
            }
        }

        // 3. Create and show popup list form
        private void ShowPopupList(SAPbouiCOM.Form parentForm)
        {
            // Check if popup already exists
            SAPbouiCOM.Form oPopupForm = null;

            try
            {
                oPopupForm = Application.SBO_Application.Forms.Item("frmPopupList");
                // Form exists, no need to create
                return;
            }
            catch
            {
                // Form does not exist, create it
            }

            // Create form
            SAPbouiCOM.FormCreationParams oParams =
                (SAPbouiCOM.FormCreationParams)Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
            oParams.UniqueID = "frmPopupList";
            oParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
            oPopupForm = Application.SBO_Application.Forms.AddEx(oParams);
            oPopupForm.Title = "Select Action";

            // Add Grid
            SAPbouiCOM.Item oItem = oPopupForm.Items.Add("gridPopup", SAPbouiCOM.BoFormItemTypes.it_GRID);
            oItem.Top = 10;
            oItem.Left = 10;
            oItem.Width = 280;
            oItem.Height = 150;

            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oItem.Specific;

            // Bind DataTable
            SAPbouiCOM.DataTable dt = oPopupForm.DataSources.DataTables.Add("DTList");
            dt.Columns.Add("ActionCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10);
            dt.Columns.Add("ActionName", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);

            // Fill list
            dt.Rows.Add(3);
            dt.SetValue("ActionCode", 0, "A1");
            dt.SetValue("ActionName", 0, "Action 1");
            dt.SetValue("ActionCode", 1, "A2");
            dt.SetValue("ActionName", 1, "Action 2");
            dt.SetValue("ActionCode", 2, "A3");
            dt.SetValue("ActionName", 2, "Action 3");

            oGrid.DataTable = dt;
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

            // Handle double-click on grid row
           Application.SBO_Application.ItemEvent += (string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent) =>
            {
                BubbleEvent = true;
                if (FormUID == "frmPopupList" && pVal.ItemUID == "gridPopup"
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK && !pVal.BeforeAction)
                {
                    string actionCode = dt.GetValue("ActionCode", pVal.Row).ToString();
                    string actionName = dt.GetValue("ActionName", pVal.Row).ToString();

                    Application.SBO_Application.MessageBox($"You selected: {actionName} ({actionCode})");

                    // Close popup
                    SAPbouiCOM.Form popup = Application.SBO_Application.Forms.Item("frmPopupList");
                    popup.Close();
                }
            };
        }


        private void ExportToXml(string FormUID)
        {
            SAPbouiCOM.ProgressBar progressBar = null;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

            try
            {
                oForm.Freeze(true);
                progressBar = Application.SBO_Application.StatusBar.CreateProgressBar("Exporting data to XML...", 0, false);
                
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
                if (progressBar != null) progressBar.Stop();
            }
        }

        private void ExportToCsv(string FormUID)
        {
            SAPbouiCOM.ProgressBar progressBar = null;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

            try
            {
                oForm.Freeze(true);
                progressBar = Application.SBO_Application.StatusBar.CreateProgressBar("Exporting data to XML...", 0, false);

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
                if (progressBar != null) progressBar.Stop();
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

            // Adjust first grid
            SAPbouiCOM.Item grid1 = oForm.Items.Item("GdFind");

            grid1.Width = formWidth - 20;
            grid1.Height = 110;

            //// Adjust second grid
            //SAPbouiCOM.Item grid2 = oForm.Items.Item("GdRes");
            //grid2.Top = grid1.Top + grid1.Height + 40; // below the first grid
            //grid2.Width = formWidth - 20;
            //grid2.Height = 140;
        }

        private void SelectAllDocHandler(string FormUID, string gdId)
        {
            SAPbouiCOM.ProgressBar progressBar = null;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            try
            {
                oForm.Freeze(true);
                progressBar = Application.SBO_Application.StatusBar.CreateProgressBar("Check/Uncheck all data...", 0, false);
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("GdFind").Specific;
                DtGridFind = oGrid.DataTable; // Use the existing DataTable binding

                // Check if the current header is marked as selected
                bool selectAll = oGrid.Columns.Item("Select").TitleObject.Caption != "[X]";
                
                for (int i = 0; i < DtGridFind.Rows.Count; i++)
                {
                    DtGridFind.SetValue("Select", i, selectAll ? "Y" : "N");
                }

                if (FindListModel != null && FindListModel.Any())
                {
                    foreach (var item in FindListModel)
                    {
                        item.Selected = selectAll;
                    }
                }
                // Update header caption
                oGrid.Columns.Item("Select").TitleObject.Caption = selectAll ? "[X]" : "[ ]";

                // Ensure the column stays as checkbox
                oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;

                oGrid.Item.Refresh();
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                oForm.Freeze(false);
                if (progressBar != null) progressBar.Stop();
            }
        }

        private void SelectSingleDocHandler(string FormUID, string gdId, int gdRow)
        {
            SAPbouiCOM.ProgressBar progressBar = null;
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            try
            {
                oForm.Freeze(true);
                progressBar = Application.SBO_Application.StatusBar.CreateProgressBar("Check/Uncheck data...", 0, false);
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(gdId).Specific;

                // Get clicked row
                int row = gdRow;

                // Get checkbox value (grid stores it as string "Y"/"N" or "tYES"/"tNO")
                string val = oGrid.DataTable.GetValue("Select", row).ToString();
                bool isChecked = val == "Y" || val == "tYES" || val == "True";
                string docEntryVal = oGrid.DataTable.GetValue("DocEntry", row).ToString();
                string objTypeVal = oGrid.DataTable.GetValue("ObjType", row).ToString();
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
            finally
            {
                oForm.Freeze(false);
                if (progressBar != null) progressBar.Stop();
            }
        }

        private void CheckBoxHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            if (pVal.ItemUID == "CkInv")
            {
                SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkInv").Specific;
                bool isChecked = oCheckBox.Checked;

                // Update dictionary
                SelectedCkBox["OINV"] = isChecked;

                coretaxModel.IsARInvoice = isChecked;
                //// Ensure the list exists
                //if (coretaxModel.SelectedDoc == null)
                //{
                //    coretaxModel.SelectedDoc = new List<string>();
                //}

                //if (isChecked)
                //{
                //    if (!coretaxModel.SelectedDoc.Contains("OINV"))
                //        coretaxModel.SelectedDoc.Add("OINV");
                //}
                //else
                //{
                //    coretaxModel.SelectedDoc.Remove("OINV"); // Safe even if item doesn't exist
                //}
            }
            if (pVal.ItemUID == "CkDp")
            {
                SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkDp").Specific;
                bool isChecked = oCheckBox.Checked;
                SelectedCkBox["ODPI"] = isChecked;

                coretaxModel.IsARDownPayment = isChecked;

                //// Ensure the list exists
                //if (coretaxModel.SelectedDoc == null)
                //{
                //    coretaxModel.SelectedDoc = new List<string>();
                //}

                //if (isChecked)
                //{
                //    if (!coretaxModel.SelectedDoc.Contains("ODPI"))
                //        coretaxModel.SelectedDoc.Add("ODPI");
                //}
                //else
                //{
                //    coretaxModel.SelectedDoc.Remove("ODPI"); // Safe even if item doesn't exist
                //}
            }
            if (pVal.ItemUID == "CkCM")
            {
                SAPbouiCOM.CheckBox oCheckBox = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkCM").Specific;
                bool isChecked = oCheckBox.Checked;
                SelectedCkBox["ORIN"] = isChecked;

                coretaxModel.IsARCreditMemo = isChecked;

                //// Ensure the list exists
                //if (coretaxModel.SelectedDoc == null)
                //{
                //    coretaxModel.SelectedDoc = new List<string>();
                //}

                //if (isChecked)
                //{
                //    if (!coretaxModel.SelectedDoc.Contains("ORIN"))
                //        coretaxModel.SelectedDoc.Add("ORIN");
                //}
                //else
                //{
                //    coretaxModel.SelectedDoc.Remove("ORIN"); // Safe even if item doesn't exist
                //}
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
            // 3. Load values using DI API Recordset
            Company oCompany = Services.CompanyService.GetCompany();
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery($"SELECT Code, Name FROM {table} ORDER BY Code"); // replace with your query

            while (!rs.EoF)
            {
                string code = rs.Fields.Item("Code").Value.ToString();
                string name = rs.Fields.Item("Name").Value.ToString();

                oCombo.ValidValues.Add(code, name);
                rs.MoveNext();
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

                        coretaxModel.FromDocNum = int.Parse(docNum);
                        coretaxModel.FromDocEntry = int.Parse(docEntry);

                        // Get the form object
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

                        // Set values into your form fields (bound to User Data Sources)
                        oForm.DataSources.UserDataSources.Item("DocFromDS").ValueEx = docNum;
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

                        coretaxModel.ToDocNum = int.Parse(docNum);
                        coretaxModel.ToDocEntry = int.Parse(docEntry);

                        // Get the form object
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

                        // Set values into your form fields (bound to User Data Sources)
                        oForm.DataSources.UserDataSources.Item("DocToDS").ValueEx = docNum;
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
                    
                    coretaxModel.FromCust = code;
                    // Get the form object
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

                    // Set values into your form fields (bound to User Data Sources)
                    oForm.DataSources.UserDataSources.Item("CustFromDS").ValueEx = code;
                }
            }
            if (oCFLEvent.ChooseFromListUID == "CflCustTo")
            {
                SAPbouiCOM.DataTable oDataTable = oCFLEvent.SelectedObjects;

                if (oDataTable != null && oDataTable.Rows.Count > 0)
                {
                    // Get values from the selected row
                    string code = oDataTable.GetValue("CardCode", 0).ToString();

                    coretaxModel.ToCust = code;

                    // Get the form object
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

                    // Set values into your form fields (bound to User Data Sources)
                    oForm.DataSources.UserDataSources.Item("CustToDS").ValueEx = code;
                }
            }
        }

        private void BtnFindHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            SAPbouiCOM.ProgressBar progressBar = null;
            try
            {
                oForm.Freeze(true);
                progressBar = Application.SBO_Application.StatusBar.CreateProgressBar("Retrieving data...", 0, false);
                progressBar.Text = "Retrieving data...";
                FindListModel.Clear();

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

                if (ItemIsExists(oForm, "CkAllDt") && oForm.Items.Item("CkAllDt").Enabled)
                {
                    SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllDt").Specific;
                    allDt = oCk.Checked;
                }
                if (ItemIsExists(oForm, "CkAllDoc") && oForm.Items.Item("CkAllDoc").Enabled)
                {
                    SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllDoc").Specific;
                    allDoc = oCk.Checked;
                }
                if (ItemIsExists(oForm, "CkAllCust") && oForm.Items.Item("CkAllCust").Enabled)
                {
                    SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllCust").Specific;
                    allCust = oCk.Checked;
                }
                if (ItemIsExists(oForm, "CkAllBr") && oForm.Items.Item("CkAllBr").Enabled)
                {
                    SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllBr").Specific;
                    allBranch = oCk.Checked;
                }
                if (ItemIsExists(oForm, "CkAllOtl") && oForm.Items.Item("CkAllOtl").Enabled)
                {
                    SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllOtl").Specific;
                    allOutlet = oCk.Checked;
                }

                if (!allDt)
                {
                    if (ItemIsExists(oForm, "TFromDt") && oForm.Items.Item("TFromDt").Enabled)
                    {
                        SAPbouiCOM.EditText oDtFrom = (SAPbouiCOM.EditText)oForm.Items.Item("TFromDt").Specific;
                        if (DateTime.TryParseExact(oDtFrom.Value, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDtFrom))
                        {
                            dtFrom = parsedDtFrom.ToString("yyyy-MM-dd");
                        }
                    }
                    if (ItemIsExists(oForm, "TToDt") && oForm.Items.Item("TToDt").Enabled)
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
                    if (ItemIsExists(oForm, "TFromDoc") && oForm.Items.Item("TFromDoc").Enabled)
                    {
                        docFrom = coretaxModel.FromDocEntry != 0 ? coretaxModel.FromDocEntry.ToString() : string.Empty;
                        //SAPbouiCOM.EditText oDocFrom = (SAPbouiCOM.EditText)oForm.Items.Item("TFromDoc").Specific;
                        //if (!string.IsNullOrEmpty(oDocFrom.Value.Trim()))
                        //{
                        //    docFrom = oDocFrom.Value.Trim();
                        //}
                    }
                    if (ItemIsExists(oForm, "TToDoc") && oForm.Items.Item("TToDoc").Enabled)
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
                    if (ItemIsExists(oForm, "TCustFrom") && oForm.Items.Item("TCustFrom").Enabled)
                    {
                        SAPbouiCOM.EditText oCustFrom = (SAPbouiCOM.EditText)oForm.Items.Item("TCustFrom").Specific;
                        if (!string.IsNullOrEmpty(oCustFrom.Value.Trim()))
                        {
                            custFrom = oCustFrom.Value.Trim();
                        }
                    }
                    if (ItemIsExists(oForm, "TCustTo") && oForm.Items.Item("TCustTo").Enabled)
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
                    if (ItemIsExists(oForm, "CbBrFrom") && oForm.Items.Item("CbBrFrom").Enabled)
                    {
                        SAPbouiCOM.ComboBox oBrFrom = (SAPbouiCOM.ComboBox)oForm.Items.Item("CbBrFrom").Specific;
                        if (!string.IsNullOrEmpty(oBrFrom.Value.Trim()))
                        {
                            branchFrom = oBrFrom.Value.Trim();
                        }
                    }
                    if (ItemIsExists(oForm, "CbBrTo") && oForm.Items.Item("CbBrTo").Enabled)
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
                    if (ItemIsExists(oForm, "CbOtlFrom") && oForm.Items.Item("CbOtlFrom").Enabled)
                    {
                        SAPbouiCOM.ComboBox oOtlFrom = (SAPbouiCOM.ComboBox)oForm.Items.Item("CbOtlFrom").Specific;
                        if (!string.IsNullOrEmpty(oOtlFrom.Value.Trim()))
                        {
                            outFrom = oOtlFrom.Value.Trim();
                        }
                    }
                    if (ItemIsExists(oForm, "CbOtlTo") && oForm.Items.Item("CbOtlTo").Enabled)
                    {
                        SAPbouiCOM.ComboBox oOtlTo = (SAPbouiCOM.ComboBox)oForm.Items.Item("CbOtlTo").Specific;
                        if (!string.IsNullOrEmpty(oOtlTo.Value.Trim()))
                        {
                            outTo = oOtlTo.Value.Trim();
                        }
                    }
                }

                SAPbobsCOM.Company oCompany = Services.CompanyService.GetCompany();
                List<string> docList = SelectedCkBox.Where((ck) => ck.Value == true).Select((c) => c.Key).ToList();
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = $@"EXEC [dbo].[T2_SP_CORETAX_HEADER] 
    @selected_doc = '{string.Join(",", docList)}'";

                List<string> filters = new List<string>();

                if (!string.IsNullOrEmpty(dtFrom))
                {
                    filters.Add($"@from_date = '{dtFrom}'");
                }
                if (!string.IsNullOrEmpty(dtTo))
                {
                    filters.Add($"@to_date = '{dtTo}'");
                }
                if (!string.IsNullOrEmpty(docFrom))
                {
                    filters.Add($"@from_docentry = {docFrom}");
                }
                if (!string.IsNullOrEmpty(docTo))
                {
                    filters.Add($"@to_docentry = {docTo}");
                }
                if (!string.IsNullOrEmpty(custFrom))
                {
                    filters.Add($"@from_cust = '{custFrom}'");
                }
                if (!string.IsNullOrEmpty(custTo))
                {
                    filters.Add($"@to_cust = '{custTo}'");
                }
                if (!string.IsNullOrEmpty(branchFrom))
                {
                    filters.Add($"@from_branch = {branchFrom}");
                }
                if (!string.IsNullOrEmpty(branchTo))
                {
                    filters.Add($"@to_branch = {branchTo}");
                }
                if (!string.IsNullOrEmpty(outFrom))
                {
                    filters.Add($"@from_outlet = {outFrom}");
                }
                if (!string.IsNullOrEmpty(outTo))
                {
                    filters.Add($"@to_outlet = {outTo}");
                }

                // join filters with commas
                if (filters.Count > 0)
                {
                    query += ", " + string.Join(", ", filters);
                }

                rs.DoQuery(query);

                while (!rs.EoF)
                {
                    var model = new FilterDataModel
                    {
                        DocEntry = rs.Fields.Item("DocEntry").Value?.ToString(),
                        DocNo = rs.Fields.Item("NoDocument").Value?.ToString(),
                        CardCode = rs.Fields.Item("BPCode").Value?.ToString(),
                        CardName = rs.Fields.Item("BPName").Value?.ToString(),
                        ObjType = rs.Fields.Item("ObjectType").Value?.ToString(),
                        ObjName = rs.Fields.Item("ObjectName").Value?.ToString(),
                        PostDate = rs.Fields.Item("InvDate").Value?.ToString(),
                        Branch = rs.Fields.Item("Branch").Value?.ToString(),
                        Outlet = rs.Fields.Item("Outlet").Value?.ToString(),
                        Selected = false,
                    };

                    FindListModel.Add(model);
                    rs.MoveNext();
                }

                SAPbouiCOM.Item mItem = oForm.Items.Item("GdFind");
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)mItem.Specific;
                if (!DtIsExists(oForm, "DT_FIND"))
                {
                    DtGridFind = oForm.DataSources.DataTables.Add("DT_FIND");
                    DtGridFind.Clear();
                }
                else
                {
                    DtGridFind = oForm.DataSources.DataTables.Item("DT_FIND");
                    DtGridFind.Clear();
                }

                DtGridFind.Columns.Add("No", SAPbouiCOM.BoFieldsType.ft_Integer, 11);
                DtGridFind.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Integer, 11);
                DtGridFind.Columns.Add("Document No", SAPbouiCOM.BoFieldsType.ft_Integer, 11);
                DtGridFind.Columns.Add("Customer Code", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                DtGridFind.Columns.Add("Customer Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 200);
                DtGridFind.Columns.Add("ObjType", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                DtGridFind.Columns.Add("Doc. Type", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                DtGridFind.Columns.Add("Posting Date", SAPbouiCOM.BoFieldsType.ft_Date);
                DtGridFind.Columns.Add("Branch", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                DtGridFind.Columns.Add("Outlet", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                DtGridFind.Columns.Add("Select", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 1);

                for (int i = 0; i < FindListModel.Count; i++)
                {
                    var item = FindListModel[i];
                    int row = DtGridFind.Rows.Count;   // ambil indeks berikutnya
                    DtGridFind.Rows.Add();
                    DtGridFind.SetValue("No", row, i+1);
                    DtGridFind.SetValue("DocEntry", row, item.DocEntry);
                    DtGridFind.SetValue("Document No", row, item.DocNo);
                    DtGridFind.SetValue("Customer Code", row, item.CardCode);
                    DtGridFind.SetValue("Customer Name", row, item.CardName);
                    DtGridFind.SetValue("ObjType", row, item.ObjType);
                    DtGridFind.SetValue("Doc. Type", row, item.ObjName);
                    DtGridFind.SetValue("Posting Date", row, DateTime.Parse(item.PostDate));
                    DtGridFind.SetValue("Branch", row, item.Branch);
                    DtGridFind.SetValue("Outlet", row, item.Outlet);
                    DtGridFind.SetValue("Select", row, "N");

                }
                int white = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

                oGrid.DataTable = DtGridFind;
                oGrid.Columns.Item("Select").Type = SAPbouiCOM.BoGridColumnType.gct_CheckBox;
                oGrid.Columns.Item("Select").TitleObject.Caption = "[ ]"; // simulate header checkbox
                
                for (int i = 0; i < oGrid.Columns.Count; i++)
                {
                    var col = oGrid.Columns.Item(i);

                    // Untuk SAP B1 ≥ 9.2: properti ada di TitleObject
                    if (col.UniqueID != "Select")
                    {
                        col.TitleObject.Sortable = true;
                        col.Editable = false;
                    }
                    if (col.UniqueID == "DocEntry" || col.UniqueID == "ObjType")
                    {
                        col.Visible = false;
                    }
                    // Jika Anda memakai versi lama dan TitleObject.Sortable belum ada,
                    // gunakan:  col.Sortable = true;
                }
                for (int i = 0; i < oGrid.Rows.Count; i++)
                {
                    oGrid.CommonSetting.SetRowBackColor(i + 1, white);
                }

                oGrid.AutoResizeColumns();
                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                oForm.Freeze(false);
                if (progressBar != null) progressBar.Stop();
                Application.SBO_Application.SetStatusBarMessage("Success", SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
        }

        private void BtnGenHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            SAPbouiCOM.ProgressBar progressBar = null;
            try
            {
                oForm.Freeze(true);
                progressBar = Application.SBO_Application.StatusBar.CreateProgressBar("Retrieving data...", 0, false);
                progressBar.Text = "Generating data...";

                var filteredHeader = FindListModel.Where((f) => f.Selected).ToList();
                if (filteredHeader != null && filteredHeader.Any())
                {
                    SAPbobsCOM.Company oCompany = Services.CompanyService.GetCompany();
                    SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    foreach (var item in filteredHeader)
                    {
                        string query = $@"EXEC [dbo].[T2_SP_CORETAX_GENERATE] 
    @DocEntry = '{item.DocEntry}', @ObjectType = '{item.ObjType}'";
                        rs.DoQuery(query);

                        while (!rs.EoF)
                        {
                            var model = new InvoiceDataModel
                            {
                                //No = rs.Fields.Item("No").Value?.ToString(),
                                TIN = rs.Fields.Item("TIN").Value?.ToString(),
                                DocEntry = rs.Fields.Item("DocEntry").Value?.ToString(),
                                LineNum = rs.Fields.Item("LineNum").Value?.ToString(),
                                InvDate = rs.Fields.Item("InvDate").Value?.ToString(),
                                NoDocument = rs.Fields.Item("NoDocument").Value?.ToString(),
                                ObjectType = rs.Fields.Item("ObjectType").Value?.ToString(),
                                ObjectName = rs.Fields.Item("ObjectName").Value?.ToString(),
                                BPCode = rs.Fields.Item("BPCode").Value?.ToString(),
                                BPName = rs.Fields.Item("BPName").Value?.ToString(),
                                SellerIDTKU = rs.Fields.Item("SellerIDTKU").Value?.ToString(),
                                BuyerDocument = rs.Fields.Item("BuyerDocument").Value?.ToString(),
                                NomorNPWP = rs.Fields.Item("NomorNPWP").Value?.ToString(),
                                NPWPName = rs.Fields.Item("NPWPName").Value?.ToString(),
                                NPWPAddress = rs.Fields.Item("NPWPAddress").Value?.ToString(),
                                BuyerIDTKU = rs.Fields.Item("BuyerIDTKU").Value?.ToString(),
                                ItemCode = rs.Fields.Item("ItemCode").Value?.ToString(),
                                DefItemCode = rs.Fields.Item("DefItemCode").Value?.ToString(),
                                ItemName = rs.Fields.Item("ItemName").Value?.ToString(),
                                ItemUnit = rs.Fields.Item("ItemUnit").Value?.ToString(),

                                // --- decimals (safe conversion) ---
                                ItemPrice = Convert.ToDouble(rs.Fields.Item("ItemPrice").Value ?? 0),
                                Qty = Convert.ToDouble(rs.Fields.Item("Qty").Value ?? 0),
                                TotalDisc = Convert.ToDouble(rs.Fields.Item("TotalDisc").Value ?? 0),
                                TaxBase = Convert.ToDouble(rs.Fields.Item("TaxBase").Value ?? 0),
                                OtherTaxBase = Convert.ToDouble(rs.Fields.Item("OtherTaxBase").Value ?? 0),
                                VATRate = Convert.ToDouble(rs.Fields.Item("VATRate").Value ?? 0),
                                AmountVAT = Convert.ToDouble(rs.Fields.Item("AmountVAT").Value ?? 0),
                                STLGRate = Convert.ToDouble(rs.Fields.Item("STLGRate").Value ?? 0),
                                STLG = Convert.ToDouble(rs.Fields.Item("STLG").Value ?? 0),

                                // --- strings ---
                                JenisPajak = rs.Fields.Item("JenisPajak").Value?.ToString(),
                                KetTambahan = rs.Fields.Item("KetTambahan").Value?.ToString(),
                                PajakPengganti = rs.Fields.Item("PajakPengganti").Value?.ToString(),
                                Referensi = rs.Fields.Item("Referensi").Value?.ToString(),
                                Status = rs.Fields.Item("Status").Value?.ToString(),
                                KodeDokumenPendukung = rs.Fields.Item("KodeDokumenPendukung").Value?.ToString(),
                                Branch = rs.Fields.Item("Branch").Value?.ToString(),
                                AddInfo = rs.Fields.Item("AddInfo").Value?.ToString(),
                                BuyerCountry = rs.Fields.Item("BuyerCountry").Value?.ToString(),
                                BuyerEmail = rs.Fields.Item("BuyerEmail").Value?.ToString(),
                            };

                            coretaxModel.Detail.Add(model);
                            rs.MoveNext();
                        }
                    }

                    //if (GenInvListModel != null && GenInvListModel.Any())
                    //{
                    //    for (int i = 0; i < GenInvListModel.Count; i++)
                    //    {
                    //        GenInvListModel[i].No = (i + 1).ToString();
                    //    }
                    //}
                }

                SAPbouiCOM.Item mItem = oForm.Items.Item("GdRes");
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)mItem.Specific;
                if (!DtIsExists(oForm, "DT_RES"))
                {
                    DtGridRes = oForm.DataSources.DataTables.Add("DT_RES");
                    DtGridRes.Clear();
                }
                else
                {
                    DtGridRes = oForm.DataSources.DataTables.Item("DT_RES");
                    DtGridRes.Clear();
                }

                // Define columns (matching all your fields)
                DtGridRes.Columns.Add("No", SAPbouiCOM.BoFieldsType.ft_Integer);
                DtGridRes.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Integer);
                DtGridRes.Columns.Add("LineNum", SAPbouiCOM.BoFieldsType.ft_Integer);
                DtGridRes.Columns.Add("Inv Date", SAPbouiCOM.BoFieldsType.ft_Date);
                DtGridRes.Columns.Add("No Document", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                DtGridRes.Columns.Add("Object Type", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                DtGridRes.Columns.Add("Object Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                DtGridRes.Columns.Add("BP Code", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                DtGridRes.Columns.Add("BP Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                DtGridRes.Columns.Add("Seller IDTKU", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                DtGridRes.Columns.Add("Buyer Document", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                DtGridRes.Columns.Add("Nomor NPWP", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                DtGridRes.Columns.Add("NPWP Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                DtGridRes.Columns.Add("NPWP Address", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 200);
                DtGridRes.Columns.Add("Buyer IDTKU", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                DtGridRes.Columns.Add("Item Code", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                DtGridRes.Columns.Add("ItemCode", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                DtGridRes.Columns.Add("Item Name", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 150);
                DtGridRes.Columns.Add("Item Unit", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                DtGridRes.Columns.Add("Item Price", SAPbouiCOM.BoFieldsType.ft_Price);
                DtGridRes.Columns.Add("Qty", SAPbouiCOM.BoFieldsType.ft_Quantity);
                DtGridRes.Columns.Add("Total Disc", SAPbouiCOM.BoFieldsType.ft_Price);
                DtGridRes.Columns.Add("Tax Base", SAPbouiCOM.BoFieldsType.ft_Price);
                DtGridRes.Columns.Add("Other Tax Base", SAPbouiCOM.BoFieldsType.ft_Price);
                DtGridRes.Columns.Add("VATRate", SAPbouiCOM.BoFieldsType.ft_Percent);
                DtGridRes.Columns.Add("Amount VAT", SAPbouiCOM.BoFieldsType.ft_Price);
                DtGridRes.Columns.Add("STLG Rate", SAPbouiCOM.BoFieldsType.ft_Percent);
                DtGridRes.Columns.Add("STLG", SAPbouiCOM.BoFieldsType.ft_Price);
                DtGridRes.Columns.Add("Jenis Pajak", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                DtGridRes.Columns.Add("Ket Tambahan", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 200);
                DtGridRes.Columns.Add("Pajak Pengganti", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                DtGridRes.Columns.Add("Referensi", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                DtGridRes.Columns.Add("Status", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10);
                DtGridRes.Columns.Add("Kode Dokumen Pendukung", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 50);
                DtGridRes.Columns.Add("Branch", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                DtGridRes.Columns.Add("Add Info", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                DtGridRes.Columns.Add("Buyer Country", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 10);
                DtGridRes.Columns.Add("Buyer Email", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);

                // Fill data
                for (int i = 0; i < coretaxModel.Detail.Count; i++)
                {
                    var item = coretaxModel.Detail[i];
                    int row = DtGridRes.Rows.Count;
                    DtGridRes.Rows.Add();

                    DtGridRes.SetValue("No", row, i + 1);
                    DtGridRes.SetValue("DocEntry", row, item.DocEntry);
                    DtGridRes.SetValue("LineNum", row, item.LineNum);
                    if (DateTime.TryParse(item.InvDate, out var invDate))
                        DtGridRes.SetValue("Inv Date", row, invDate);
                    DtGridRes.SetValue("No Document", row, item.NoDocument);
                    DtGridRes.SetValue("Object Type", row, item.ObjectType);
                    DtGridRes.SetValue("Object Name", row, item.ObjectName);
                    DtGridRes.SetValue("BP Code", row, item.BPCode);
                    DtGridRes.SetValue("BP Name", row, item.BPName);
                    DtGridRes.SetValue("Seller IDTKU", row, item.SellerIDTKU);
                    DtGridRes.SetValue("Buyer Document", row, item.BuyerDocument);
                    DtGridRes.SetValue("Nomor NPWP", row, item.NomorNPWP);
                    DtGridRes.SetValue("NPWP Name", row, item.NPWPName);
                    DtGridRes.SetValue("NPWP Address", row, item.NPWPAddress);
                    DtGridRes.SetValue("Buyer IDTKU", row, item.BuyerIDTKU);
                    DtGridRes.SetValue("Item Code", row, item.DefItemCode);
                    DtGridRes.SetValue("ItemCode", row, item.ItemCode);
                    DtGridRes.SetValue("Item Name", row, item.ItemName);
                    DtGridRes.SetValue("Item Unit", row, item.ItemUnit);
                    DtGridRes.SetValue("Item Price", row, item.ItemPrice);
                    DtGridRes.SetValue("Qty", row, item.Qty);
                    DtGridRes.SetValue("Total Disc", row, item.TotalDisc);
                    DtGridRes.SetValue("Tax Base", row, item.TaxBase);
                    DtGridRes.SetValue("Other Tax Base", row, item.OtherTaxBase);
                    DtGridRes.SetValue("VATRate", row, item.VATRate);
                    DtGridRes.SetValue("Amount VAT", row, item.AmountVAT);
                    DtGridRes.SetValue("STLG Rate", row, item.STLGRate);
                    DtGridRes.SetValue("STLG", row, item.STLG);
                    DtGridRes.SetValue("Jenis Pajak", row, item.JenisPajak);
                    DtGridRes.SetValue("Ket Tambahan", row, item.KetTambahan);
                    DtGridRes.SetValue("Pajak Pengganti", row, item.PajakPengganti);
                    DtGridRes.SetValue("Referensi", row, item.Referensi);
                    DtGridRes.SetValue("Status", row, item.Status);
                    DtGridRes.SetValue("Kode Dokumen Pendukung", row, item.KodeDokumenPendukung);
                    DtGridRes.SetValue("Branch", row, item.Branch);
                    DtGridRes.SetValue("Add Info", row, item.AddInfo);
                    DtGridRes.SetValue("Buyer Country", row, item.BuyerCountry);
                    DtGridRes.SetValue("Buyer Email", row, item.BuyerEmail);
                }

                int white = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

                oGrid.DataTable = DtGridRes;
                oGrid.AutoResizeColumns();
                oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;

                for (int i = 0; i < oGrid.Columns.Count; i++)
                {
                    var col = oGrid.Columns.Item(i);
                    
                    if (col.UniqueID == "DocEntry" 
                        || col.UniqueID == "LineNum"
                        || col.UniqueID == "ItemCode"
                        )
                    {
                        col.Visible = false;
                    }
                    // Jika Anda memakai versi lama dan TitleObject.Sortable belum ada,
                    // gunakan:  col.Sortable = true;
                }
                for (int i = 0; i < oGrid.Rows.Count; i++)
                {
                    oGrid.CommonSetting.SetRowBackColor(i + 1, white);
                }

                SAPbouiCOM.Button btAdd = (SAPbouiCOM.Button)oForm.Items.Item("BtAdd").Specific;
                if (coretaxModel.Detail.Any())
                {
                    btAdd.Item.Enabled = true;
                }
                else
                {
                    btAdd.Item.Enabled = false;
                }

                //SAPbouiCOM.Button btXml = (SAPbouiCOM.Button)oForm.Items.Item("BtXML").Specific;
                //SAPbouiCOM.Button btCsv = (SAPbouiCOM.Button)oForm.Items.Item("BtCSV").Specific;
                //if (GenInvListModel.Any())
                //{
                //    btXml.Item.Enabled = true;
                //    btCsv.Item.Enabled = true;
                //}
                //else
                //{
                //    btXml.Item.Enabled = false;
                //    btCsv.Item.Enabled = false;
                //}
            }
            catch (Exception e)
            {

                throw e;
            }
            finally
            {
                oForm.Freeze(false);
                if (progressBar != null) progressBar.Stop();
                Application.SBO_Application.SetStatusBarMessage("Success", SAPbouiCOM.BoMessageTime.bmt_Short, false);
            }
        }

        private void ShowFilterGroup(SAPbouiCOM.Form oForm)
        {
            oForm.ActiveItem = "TDocNum";
            if (FilterIsShow)
            {
                var countDoc = SelectedCkBox.Where((d) => d.Value).Count();
                if (!oForm.Items.Item("TFromDt").Enabled)
                {
                    oForm.Items.Item("TFromDt").Enabled = true;
                }
                if (!oForm.Items.Item("TToDt").Enabled)
                {
                    oForm.Items.Item("TToDt").Enabled = true;
                }
                if (!oForm.Items.Item("CkAllDt").Enabled)
                {
                    oForm.Items.Item("CkAllDt").Enabled = true;
                }

                if (countDoc == 1)
                {
                    var table = SelectedCkBox.Where((d) => d.Value).First().Key;
                    if (!oForm.Items.Item("TFromDoc").Enabled)
                    {
                        oForm.Items.Item("TFromDoc").Enabled = true;
                    }
                    if (!oForm.Items.Item("TToDoc").Enabled)
                    {
                        oForm.Items.Item("TToDoc").Enabled = true;
                    }
                    if (!oForm.Items.Item("CkAllDoc").Enabled)
                    {
                        oForm.Items.Item("CkAllDoc").Enabled = true;
                    }
                    SetDocumentCfl(oForm, "CflDocFrom"+table, "DocFromDS", "TFromDoc");
                    SetDocumentCfl(oForm, "CflDocTo"+table, "DocToDS", "TToDoc");
                }
                else
                {
                    if (oForm.Items.Item("TFromDoc").Enabled)
                    {
                        oForm.Items.Item("TFromDoc").Enabled = false;
                        ClearEdit(oForm, "TFromDoc");
                    }
                    if (oForm.Items.Item("TToDoc").Enabled)
                    {
                        oForm.Items.Item("TToDoc").Enabled = false;
                        ClearEdit(oForm, "TToDoc");
                    }
                    if (oForm.Items.Item("CkAllDoc").Enabled)
                    {
                        oForm.Items.Item("CkAllDoc").Enabled = false;
                        ClearCheckBox(oForm, "CkAllDoc", "CkDocDS");
                    }
                }

                if (!oForm.Items.Item("TCustFrom").Enabled)
                {
                    oForm.Items.Item("TCustFrom").Enabled = true;
                }
                if (!oForm.Items.Item("TCustTo").Enabled)
                {
                    oForm.Items.Item("TCustTo").Enabled = true;
                }
                if (!oForm.Items.Item("CkAllCust").Enabled)
                {
                    oForm.Items.Item("CkAllCust").Enabled = true;
                }

                if (!oForm.Items.Item("CbBrFrom").Enabled)
                {
                    oForm.Items.Item("CbBrFrom").Enabled = true;
                }
                if (!oForm.Items.Item("CbBrTo").Enabled)
                {
                    oForm.Items.Item("CbBrTo").Enabled = true;
                }
                if (!oForm.Items.Item("CkAllBr").Enabled)
                {
                    oForm.Items.Item("CkAllBr").Enabled = true;
                }

                if (!oForm.Items.Item("CbOtlFrom").Enabled)
                {
                    oForm.Items.Item("CbOtlFrom").Enabled = true;
                }
                if (!oForm.Items.Item("CbOtlTo").Enabled)
                {
                    oForm.Items.Item("CbOtlTo").Enabled = true;
                }
                if (!oForm.Items.Item("CkAllOtl").Enabled)
                {
                    oForm.Items.Item("CkAllOtl").Enabled = true;
                }

                SetCustomerCfl(oForm, "CflCustFrom", "CustFromDS", "TCustFrom");
                SetCustomerCfl(oForm, "CflCustTo", "CustToDS", "TCustTo");

                GetDataComboBox(oForm, "CbBrFrom", "OUBR");
                GetDataComboBox(oForm, "CbBrTo", "OUBR");

                GetDataComboBox(oForm, "CbOtlFrom", "OUDP");
                GetDataComboBox(oForm, "CbOtlTo", "OUDP");

                oForm.Items.Item("BtFind").Enabled = true;
            }
            else
            {
                if (oForm.Items.Item("TFromDt").Enabled)
                {
                    oForm.Items.Item("TFromDt").Enabled = false;
                    ClearEdit(oForm, "TFromDt");
                }
                if (oForm.Items.Item("TToDt").Enabled)
                {
                    oForm.Items.Item("TToDt").Enabled = false;
                    ClearEdit(oForm, "TToDt");
                }
                if (oForm.Items.Item("CkAllDt").Enabled)
                {
                    oForm.Items.Item("CkAllDt").Enabled = false;
                    ClearCheckBox(oForm, "CkAllDt", "CkDtDS");
                }

                if (oForm.Items.Item("TFromDoc").Enabled)
                {
                    oForm.Items.Item("TFromDoc").Enabled = false;
                    ClearEdit(oForm, "TFromDoc");
                }
                if (oForm.Items.Item("TToDoc").Enabled)
                {
                    oForm.Items.Item("TToDoc").Enabled = false;
                    ClearEdit(oForm, "TToDoc");
                }
                if (oForm.Items.Item("CkAllDoc").Enabled)
                {
                    oForm.Items.Item("CkAllDoc").Enabled = false;
                    ClearCheckBox(oForm, "CkAllDoc", "CkDocDS");
                }

                if (oForm.Items.Item("TCustFrom").Enabled)
                {
                    oForm.Items.Item("TCustFrom").Enabled = false;
                    ClearEdit(oForm, "TCustFrom");
                }
                if (oForm.Items.Item("TCustTo").Enabled)
                {
                    oForm.Items.Item("TCustTo").Enabled = false;
                    ClearEdit(oForm, "TCustTo");
                }
                if (oForm.Items.Item("CkAllCust").Enabled)
                {
                    oForm.Items.Item("CkAllCust").Enabled = false;
                    ClearCheckBox(oForm, "CkAllCust", "CkCustDS");
                }

                if (oForm.Items.Item("CbBrFrom").Enabled)
                {
                    oForm.Items.Item("CbBrFrom").Enabled = false;
                    ClearCombo(oForm, "CbBrFrom");
                }
                if (oForm.Items.Item("CbBrTo").Enabled)
                {
                    oForm.Items.Item("CbBrTo").Enabled = false;
                    ClearCombo(oForm, "CbBrTo");
                }
                if (oForm.Items.Item("CkAllBr").Enabled)
                {
                    oForm.Items.Item("CkAllBr").Enabled = false;
                    ClearCheckBox(oForm, "CkAllBr", "CkBrDS");
                }

                if (oForm.Items.Item("CbOtlFrom").Enabled)
                {
                    oForm.Items.Item("CbOtlFrom").Enabled = false;
                    ClearCombo(oForm, "CbOtlFrom");
                }
                if (oForm.Items.Item("CbOtlTo").Enabled)
                {
                    oForm.Items.Item("CbOtlTo").Enabled = false;
                    ClearCombo(oForm, "CbOtlTo");
                }
                if (oForm.Items.Item("CkAllOtl").Enabled)
                {
                    oForm.Items.Item("CkAllOtl").Enabled = false;
                    ClearCheckBox(oForm, "CkAllOtl", "CkOtlDS");
                }

                oForm.Items.Item("BtFind").Enabled = false;
            }
            ClearGrid(oForm, "GdFind", "DT_FIND");
            ClearGrid(oForm, "GdRes", "DT_RES");
            oForm.Items.Item("BtAdd").Enabled = false;
        }

        private void ClearCombo(SAPbouiCOM.Form form, string id)
        {
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
            SAPbouiCOM.EditText edit = (SAPbouiCOM.EditText)oForm.Items.Item(id).Specific;
            edit.Value = string.Empty;
        }

        private void ClearCheckBox(SAPbouiCOM.Form oForm, string id, string ds)
        {
            SAPbouiCOM.CheckBox ck = (SAPbouiCOM.CheckBox)oForm.Items.Item(id).Specific;
            ck.Checked = false;
            oForm.DataSources.UserDataSources.Item(ds).Value = "N";
        }

        private void SetBoldUnderlinedLabel(SAPbouiCOM.Form oForm, string id)
        {
            if (ItemIsExists(oForm, id))
            {
                // Get StaticText item
                SAPbouiCOM.StaticText label = (SAPbouiCOM.StaticText)oForm.Items.Item(id).Specific;

                // Bold Underline
                label.Item.TextStyle = (int)(SAPbouiCOM.BoFontStyle.fs_Bold | SAPbouiCOM.BoFontStyle.fs_Underline);
            }
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

            // ✅ Check if the item exists before binding
            if (ItemIsExists(oForm, txtId))
            {
                var oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(txtId).Specific;
                oEdit.DataBind.SetBound(true, "", ds);
                oEdit.ChooseFromListUID = id;      // ✅ use your parameter `id` instead of hardcoded "2"
                oEdit.ChooseFromListAlias = "CardCode";
            }
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

                // ✅ Check if the item exists before binding
                if (ItemIsExists(oForm, txtId))
                {
                    var oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(txtId).Specific;
                    oEdit.DataBind.SetBound(true, "", ds);
                    oEdit.ChooseFromListUID = id;      // ✅ use your parameter `id` instead of hardcoded "2"
                                                       //oEdit.ChooseFromListAlias = "CardCode";
                }
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

            // ✅ Bind EditText to DataSource and CFL
            if (ItemIsExists(oForm, txtId))
            {
                var oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(txtId).Specific;
                oEdit.DataBind.SetBound(true, "", ds);
                oEdit.ChooseFromListUID = id;
                //oEdit.ChooseFromListAlias = "Code"; // or "BPLName" if you want the name
            }
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
        
    }
}
