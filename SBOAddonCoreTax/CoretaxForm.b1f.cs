using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;

namespace SBOAddonCoreTax
{
    [FormAttribute("SBOAddonCoreTax.CoretaxForm", "CoretaxForm.b1f")]
    class CoretaxForm : UserFormBase
    {
        private SAPbouiCOM.CheckBox CkInv;
        private SAPbouiCOM.CheckBox CkDp;
        private SAPbouiCOM.CheckBox CkCM;
        private SAPbouiCOM.StaticText lDate;
        private SAPbouiCOM.EditText TFromDt;
        private SAPbouiCOM.EditText TToDt;
        private SAPbouiCOM.Grid GRes;
        private SAPbouiCOM.Button BtAdd;
        private SAPbouiCOM.Button BtCancel;
        private SAPbouiCOM.StaticText lDocF;
        private SAPbouiCOM.EditText TFromDoc;
        private SAPbouiCOM.EditText TToDoc;
        private SAPbouiCOM.StaticText lCust;
        private SAPbouiCOM.EditText TCustFrom;
        private SAPbouiCOM.EditText TCustTo;
        private SAPbouiCOM.Button BtFind;
        private SAPbouiCOM.Button BtGen;
        private SAPbouiCOM.StaticText lDocNo;
        private SAPbouiCOM.EditText TDocNum;
        public CoretaxForm()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.lDocNo = ((SAPbouiCOM.StaticText)(this.GetItem("lDocNum").Specific));
            this.TDocNum = ((SAPbouiCOM.EditText)(this.GetItem("TDocNum").Specific));
            this.CkInv = ((SAPbouiCOM.CheckBox)(this.GetItem("CkInv").Specific));
            this.CkDp = ((SAPbouiCOM.CheckBox)(this.GetItem("CkDp").Specific));
            this.CkCM = ((SAPbouiCOM.CheckBox)(this.GetItem("CkCM").Specific));
            this.lDate = ((SAPbouiCOM.StaticText)(this.GetItem("lDate").Specific));
            this.TFromDt = ((SAPbouiCOM.EditText)(this.GetItem("TFromDt").Specific));
            this.TToDt = ((SAPbouiCOM.EditText)(this.GetItem("TToDt").Specific));
            this.BtFind = ((SAPbouiCOM.Button)(this.GetItem("BtFind").Specific));
            this.BtGen = ((SAPbouiCOM.Button)(this.GetItem("BtGen").Specific));
            this.GRes = ((SAPbouiCOM.Grid)(this.GetItem("GdRes").Specific));
            this.BtAdd = ((SAPbouiCOM.Button)(this.GetItem("BtAdd").Specific));
            this.BtCancel = ((SAPbouiCOM.Button)(this.GetItem("BtCancel").Specific));
            this.lDocF = ((SAPbouiCOM.StaticText)(this.GetItem("lDocF").Specific));
            this.TFromDoc = ((SAPbouiCOM.EditText)(this.GetItem("TFromDoc").Specific));
            this.TToDoc = ((SAPbouiCOM.EditText)(this.GetItem("TToDoc").Specific));
            this.lCust = ((SAPbouiCOM.StaticText)(this.GetItem("lCust").Specific));
            this.TCustFrom = ((SAPbouiCOM.EditText)(this.GetItem("TCustFrom").Specific));
            this.TCustTo = ((SAPbouiCOM.EditText)(this.GetItem("TCustTo").Specific));
            this.TCustTo.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.TCustTo_KeyDownAfter);
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("GdFind").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("lToDt").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("lToDoc").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("lToCust").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("lFromDt").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("lFromDoc").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("lFromCust").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("lDisp").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("lParam").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("lBranch").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("lFromBr").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("lToBr").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("lOtl").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("lFromOtl").Specific));
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("lToOtl").Specific));
            this.StaticText14 = ((SAPbouiCOM.StaticText)(this.GetItem("lStatus").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("TStatus").Specific));
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("CkAllDt").Specific));
            this.CheckBox1 = ((SAPbouiCOM.CheckBox)(this.GetItem("CkAllDoc").Specific));
            this.CheckBox2 = ((SAPbouiCOM.CheckBox)(this.GetItem("CkAllCust").Specific));
            this.CheckBox3 = ((SAPbouiCOM.CheckBox)(this.GetItem("CkAllBr").Specific));
            this.CheckBox4 = ((SAPbouiCOM.CheckBox)(this.GetItem("CkAllOtl").Specific));
            this.CbBrFrom = ((SAPbouiCOM.ComboBox)(this.GetItem("CbBrFrom").Specific));
            this.CbBrTo = ((SAPbouiCOM.ComboBox)(this.GetItem("CbBrTo").Specific));
            this.CbOtlFrom = ((SAPbouiCOM.ComboBox)(this.GetItem("CbOtlFrom").Specific));
            this.CbOtlTo = ((SAPbouiCOM.ComboBox)(this.GetItem("CbOtlTo").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("BtXML").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("BtCSV").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);

        }

        private void OnCustomInitialize()
        {
            
        }

        private SAPbouiCOM.Grid Grid0;


        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;

        private void TCustTo_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
           
        }

        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.StaticText StaticText11;
        private SAPbouiCOM.StaticText StaticText12;
        private SAPbouiCOM.StaticText StaticText13;
        private SAPbouiCOM.StaticText StaticText14;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.CheckBox CheckBox0;
        private SAPbouiCOM.CheckBox CheckBox1;
        private SAPbouiCOM.CheckBox CheckBox2;
        private SAPbouiCOM.CheckBox CheckBox3;
        private SAPbouiCOM.CheckBox CheckBox4;
        private SAPbouiCOM.ComboBox CbBrFrom;
        private SAPbouiCOM.ComboBox CbBrTo;
        private SAPbouiCOM.ComboBox CbOtlFrom;
        private SAPbouiCOM.ComboBox CbOtlTo;

        public void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
        }

        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Button Button1;
    }
}