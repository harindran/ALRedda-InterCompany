using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using General.Common;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;

namespace ALRedda.Business_Objects
{
    [FormAttribute("170", "Business_Objects/Incpay.b1f")]
    class Incpay : SystemFormBase
    {
        public SAPbouiCOM.Form oForm;
        private SAPbobsCOM.Company objAnothercompany;
        private string BPCodeCustomer = "";
        private string BPCodeVendor = "";
        private string Othercompany = "";

        public Incpay()
        {
        }

        private StaticText StaticText0;
        private StaticText StaticText1;
        private StaticText StaticText2;        
        private ComboBox ComboBox0;
        private EditText EditText0;        
        private ComboBox ComboBox1;
        private Folder Folder0;
        private Matrix Matrix1;

        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_0").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_2").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_3").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("Item_5").Specific));
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Transpage").Specific));
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_7").Specific));
            this.Matrix1.ClickAfter += new SAPbouiCOM._IMatrixEvents_ClickAfterEventHandler(this.Matrix1_ClickAfter);
            this.ComboBox1.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox1_ComboSelectAfter);
            this.ComboBox0.ComboSelectAfter += new SAPbouiCOM._IComboBoxEvents_ComboSelectAfterEventHandler(this.ComboBox0_ComboSelectAfter);
            this.EditText0.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.EditText0_LostFocusAfter);
            this.Folder0.ClickBefore += new SAPbouiCOM._IFolderEvents_ClickBeforeEventHandler(this.Folder0_ClickBefore);
            this.OnCustomInitialize();

        }

        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.DataAddAfter += new SAPbouiCOM.Framework.FormBase.DataAddAfterHandler(this.Form_DataAddAfter);
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);
        }
        
        private void OnCustomInitialize()
        {
            this.Folder0.GroupWith("253000196");
            Matrix1.Columns.Item("#").Visible = false;
        }
        
        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm = clsModule.objaddon.objapplication.Forms.GetForm("170", pVal.FormTypeCount);
            string strSQL = "SELECT * FROM SCHEMAS WHERE SCHEMA_OWNER='SYSTEM' and SCHEMA_NAME not in('" + clsModule.objaddon.objcompany.CompanyDB + "','SBOCOMMON');";
            clsModule.objaddon.objglobalmethods.WriteErrorLog(strSQL);
            oForm.Items.Item("37").ToPane = 21;
            oForm.Items.Item("13").ToPane = 21;
            clsModule.objaddon.objglobalmethods.Load_Combo(oForm.UniqueID, this.ComboBox0, strSQL);
            oForm.Items.Item("Item_4").Visible = oForm.Items.Item("1320002037").Visible;
            oForm.Items.Item("Item_5").Visible = oForm.Items.Item("1320002037").Visible;       
        }
       
        private void Folder0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            oForm.PaneLevel = 21;
        }

        private void Loadtotvalue(SAPbouiCOM.SBOItemEventArg pVal)
        {
            decimal totvalue = 0;
            for (int i = 0; i < Matrix1.RowCount; i++)
            {
                if (((SAPbouiCOM.CheckBox)Matrix1.Columns.Item("Select").Cells.Item(i + 1).Specific).Checked)
                    totvalue += Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix1.Columns.Item("TotPay").Cells.Item(i + 1).Specific).Value.ToString());
            }
            ((SAPbouiCOM.CheckBox)oForm.Items.Item("37").Specific).Checked = true;
            ((SAPbouiCOM.EditText)oForm.Items.Item("13").Specific).Value = totvalue.ToString();
        }
        
        private void ComboBox0_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (string.IsNullOrEmpty(ComboBox0.Selected.Value)) return;

            if (oForm.Items.Item("Item_5").Visible)
            {
                string strSQL = "SELECT \"BPLName\",\"BPLId\"  FROM " + ComboBox0.Selected.Value + ".OBPL o ;";
                clsModule.objaddon.objglobalmethods.Load_Combo(oForm.UniqueID, this.ComboBox1, strSQL);
            }
            if (Othercompany != this.ComboBox0.Value )
            {                
                Othercompany = this.ComboBox0.Value;
            }
            Transaction.addLoadMatrixCol(Matrix1, ComboBox0.Selected.Value);
        }

        private void ComboBox1_ComboSelectAfter(object sboObject, SBOItemEventArg pVal)
        {
            Loaddata();
        }

        private void Matrix1_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ColUID == "Select")
            {
                Loadtotvalue(pVal);
            }
        }

        private void Form_DataAddAfter(ref BusinessObjectInfo pVal)
        {
            if (pVal.ActionSuccess)
            {
                Transaction.anotherCompany(this.ComboBox0.Value, out objAnothercompany, out BPCodeCustomer, out BPCodeVendor);
                PostVoucher();
                Transaction.saveTransaction(oForm, Matrix1);
                Matrix1.Clear();
            }
        }

        private bool PostVoucher()
        {
            SAPbobsCOM.JournalEntries oJV;
            oJV = (SAPbobsCOM.JournalEntries)objAnothercompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            double totvalue = 0;
            int branch = 0;
            int col = 0;

            for (int i = 0; i < Matrix1.RowCount; i++)
            {
                bool ss2 = ((SAPbouiCOM.CheckBox)Matrix1.Columns.Item("Select").Cells.Item(i + 1).Specific).Checked;
                if (ss2)
                {
                    oJV.Lines.ShortName = EditText0.Value;
                    oJV.Lines.BPLID = Convert.ToInt32(((SAPbouiCOM.EditText)Matrix1.Columns.Item("BPLId").Cells.Item(i + 1).Specific).Value.ToString());
                    oJV.Lines.Credit = Convert.ToDouble(((SAPbouiCOM.EditText)Matrix1.Columns.Item("TotPay").Cells.Item(i + 1).Specific).Value.ToString());
                    do
                    {
                        string column = "Dim" + (col + 1).ToString();
                        if (clsModule.objaddon.objglobalmethods.GetColumnindex(Matrix1, column) == -1) break;

                        switch (col + 1)
                        {
                            case 1:
                                oJV.Lines.CostingCode = ((SAPbouiCOM.EditText)Matrix1.Columns.Item(column).Cells.Item(i + 1).Specific).Value.ToString();
                                break;
                            case 2:
                                oJV.Lines.CostingCode2 = ((SAPbouiCOM.EditText)Matrix1.Columns.Item(column).Cells.Item(i + 1).Specific).Value.ToString();
                                break;
                            case 3:
                                oJV.Lines.CostingCode3 = ((SAPbouiCOM.EditText)Matrix1.Columns.Item(column).Cells.Item(i + 1).Specific).Value.ToString();
                                break;
                            case 4:
                                oJV.Lines.CostingCode4 = ((SAPbouiCOM.EditText)Matrix1.Columns.Item(column).Cells.Item(i + 1).Specific).Value.ToString();
                                break;
                            case 5:
                                oJV.Lines.CostingCode5 = ((SAPbouiCOM.EditText)Matrix1.Columns.Item(column).Cells.Item(i + 1).Specific).Value.ToString();
                                break;
                        }
                        col++;
                    } while (true);
                    oJV.Lines.Add();
                    totvalue += Convert.ToDouble(((SAPbouiCOM.EditText)Matrix1.Columns.Item("TotPay").Cells.Item(i + 1).Specific).Value.ToString());
                    branch = Convert.ToInt32(((SAPbouiCOM.EditText)Matrix1.Columns.Item("BPLId").Cells.Item(i + 1).Specific).Value.ToString());
                }
            }
            string BPcode = "";
            switch (oForm.DataSources.DBDataSources.Item("ORCT").GetValue("DocType", 0))
            {
                case "C":
                    BPcode = BPCodeCustomer;
                    break;
                case "S":
                    BPcode = BPCodeVendor;
                    break;
            }
            oJV.Lines.ShortName = BPcode;
            oJV.Lines.BPLID = branch;
            oJV.Lines.Debit = totvalue;
            oJV.Lines.Add();

            int iErrCode = oJV.Add();
            string strerr = "";
            if (iErrCode != 0)
            {
                objAnothercompany.GetLastError(out iErrCode, out strerr);                
                clsModule.objaddon.objapplication.MessageBox(strerr);
            }
            else
            {
                string udfValue = objAnothercompany.GetNewObjectKey();
                Transaction.savePayUDF(Convert.ToInt32(oForm.DataSources.DBDataSources.Item("ORCT").GetValue("DocEntry", 0)), "U_JouEnt", udfValue, BoObjectTypes.oIncomingPayments);
                Transaction.reconcile(objAnothercompany, oForm, Matrix1, EditText0.Value, Convert.ToInt32(udfValue),BoObjectTypes.oIncomingPayments);               
            }
            return true;
        }

        private void Form_DataLoadAfter(ref BusinessObjectInfo pVal)
        {
            Transaction.loadSaveData(oForm, Matrix1);
            oForm.Mode = BoFormMode.fm_OK_MODE;
        }

        private void EditText0_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (!oForm.Items.Item("1320002037").Visible)
            {
                Loaddata();
            }
        }

        private void Loaddata()
        {
            try
            {
                oForm.Freeze(true);
                oForm.Items.Item("Transpage").Click();
                string lstrquery = "";
                lstrquery += " SELECT 'N' as \"Selected\",o.\"CreatedBy\" AS \"DocEntry\" , o.\"BaseRef\" AS \"DocumentNo\", ";
                lstrquery += " CASE WHEN j.\"TransType\" =18 THEN  'PU' ";
                lstrquery += " WHEN j.\"TransType\" =30 THEN 'JE' ";
                lstrquery += " WHEN j.\"TransType\" =13 THEN 'IN' ";
                lstrquery += " WHEN j.\"TransType\" =14 THEN 'CN' ";
                lstrquery += " WHEN j.\"TransType\" =24 THEN 'RC' ";
                lstrquery += " WHEN j.\"TransType\" =203 THEN 'JE' ELSE ''end ";
                lstrquery += " AS \"DocType\",o.\"RefDate\" AS  \"DocDate\",j.\"BPLId\", ";
                lstrquery += " DAYS_BETWEEN(o.\"RefDate\",CURRENT_DATE) AS \"OverDue Days\", ";
                lstrquery += " (SELECT Top 1 \"SysCurrncy\"  FROM OADM o) || ' ' || CAST ((j.\"SYSCred\"-j.\"SYSDeb\") *-1 AS Varchar(100))  AS \"SysAmount\", ";
                lstrquery += " o.\"TransCurr\" || ' ' || CAST((j.\"FCCredit\" - j.\"FCDebit\")*-1 AS Varchar(100)) AS \"FcAmount\", ";
                lstrquery += " (j.\"BalScCred\"-j.\"BalScDeb\")*-1 AS \"Total Amount\", ";
                lstrquery += "  (j.\"BalScCred\"-j.\"BalScDeb\")*-1 AS \"Total Payment\" ";
                lstrquery += " FROM " + ComboBox0.Selected.Value + ".OJDT o ";
                lstrquery += " INNER JOIN  " + ComboBox0.Selected.Value + ".JDT1 j ON o.\"TransId\" =j.\"TransId\" ";
                lstrquery += " WHERE o.\"BtfStatus\" ='O' AND j.\"TransType\" in(13,30,14,203,24) ";
                lstrquery += " AND (j.\"BalScCred\"-j.\"BalScDeb\") <>0 ";
                lstrquery += " and j.\"ShortName\" ='" + EditText0.Value + "' ";
                if (oForm.Items.Item("Item_5").Visible)
                    lstrquery += " AND j.\"BPLName\"  ='" + ComboBox1.Selected.Value + "' ";
                lstrquery += " Order by \"DocDate\",j.\"BPLId\" ";
                clsModule.objaddon.objglobalmethods.WriteErrorLog(lstrquery);
                SAPbobsCOM.Recordset rc = clsModule.objaddon.objglobalmethods.GetmultipleRS(lstrquery);
                Matrix1.Clear();

                if (rc.RecordCount > 0)
                {

                    Matrix1.Columns.Item("Select").Editable = true;
                    Matrix1.Columns.Item("TotPay").Editable = true;
                    for (int i = 0; i < rc.RecordCount; i++)
                    {
                        Matrix1.AddRow();
                        ((SAPbouiCOM.CheckBox)Matrix1.Columns.Item("Select").Cells.Item(Matrix1.VisualRowCount).Specific).Checked = rc.Fields.Item("Selected").Value.ToString() != "N";
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("DocEntry").Cells.Item(Matrix1.VisualRowCount).Specific).String = rc.Fields.Item("DocEntry").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("DocNo").Cells.Item(Matrix1.VisualRowCount).Specific).String = rc.Fields.Item("DocumentNo").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("DocType").Cells.Item(Matrix1.VisualRowCount).Specific).String = rc.Fields.Item("DocType").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("BPLId").Cells.Item(Matrix1.VisualRowCount).Specific).String = rc.Fields.Item("BPLId").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("DocDate").Cells.Item(Matrix1.VisualRowCount).Specific).String = clsModule.objaddon.objglobalmethods.Getdateformat(rc.Fields.Item("DocDate").Value.ToString());
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("OverDD").Cells.Item(Matrix1.VisualRowCount).Specific).String = rc.Fields.Item("OverDue Days").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("SysDocval").Cells.Item(Matrix1.VisualRowCount).Specific).String = rc.Fields.Item("SysAmount").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("FCDocAmt").Cells.Item(Matrix1.VisualRowCount).Specific).String = rc.Fields.Item("FcAmount").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("TotAmt").Cells.Item(Matrix1.VisualRowCount).Specific).String = rc.Fields.Item("Total Amount").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("TotPay").Cells.Item(Matrix1.VisualRowCount).Specific).String = rc.Fields.Item("Total Payment").Value.ToString();
                        rc.MoveNext();
                    }
                }
                SAPbouiCOM.Matrix contgrid = (SAPbouiCOM.Matrix)oForm.Items.Item("20").Specific;
                if (oForm.Mode == BoFormMode.fm_ADD_MODE)
                    Matrix1.Columns.Item("Select").Visible = true;
            }

            catch (Exception ex)
            {


            }
            finally
            {
                oForm.Freeze(false);
            }
        }

    }
}
