using General.Common;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ALRedda.Business_Objects
{


    [FormAttribute("392_NoUse", "Business_Objects/JEPay.b1f")]
    public class JEPay : SystemFormBase
    {
        public SAPbouiCOM.Form oForm;
        private int Currentrow = 0;
        private SAPbobsCOM.Company objAnothercompany;
        private string BPCodeCustomer = "";
        private string BPCodeVendor = "";

        public JEPay()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("76").Specific));
            this.Matrix0.ClickBefore += new SAPbouiCOM._IMatrixEvents_ClickBeforeEventHandler(this.Matrix0_ClickBefore);
            this.Matrix0.GotFocusAfter += new SAPbouiCOM._IMatrixEvents_GotFocusAfterEventHandler(this.Matrix0_GotFocusAfter);
            this.Matrix0.KeyDownBefore += new SAPbouiCOM._IMatrixEvents_KeyDownBeforeEventHandler(this.Matrix0_KeyDownBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.DataAddAfter += new SAPbouiCOM.Framework.FormBase.DataAddAfterHandler(this.Form_DataAddAfter);
            this.ActivateAfter += new SAPbouiCOM.Framework.FormBase.ActivateAfterHandler(this.Form_ActivateAfter);


        }

        private SAPbouiCOM.Matrix Matrix0;



        private void OnCustomInitialize()
        {
            try
            {

            }
            catch (Exception ex)
            {

                throw;
            }

        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm = clsModule.objaddon.objapplication.Forms.GetForm("392", pVal.FormTypeCount);
        }

        private void Matrix0_KeyDownBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.ColUID == "U_ODBAccNo" && pVal.CharPressed == 9)
            {

                if (!string.IsNullOrEmpty(((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBAccNo").Cells.Item(pVal.Row).Specific).Value)) return;
                if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBNames").Cells.Item(pVal.Row).Specific).Value)) return;
                choose choose = new choose();
                choose.Retval += Choose_Retval;

                if (pVal.Modifiers != SAPbouiCOM.BoModifiersEnum.mt_CTRL)
                {
                    choose.lstrquery = "SELECT  \"AcctName\" as \"Name\" ,\"AcctCode\" as \"code\", \"AcctCode\" as \"CtrlCode\" FROM " + ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBNames").Cells.Item(pVal.Row).Specific).Value.ToString() + ".OACT  where \"LocManTran\" ='N' and  \"Postable\" ='Y'  AND \"FrozenFor\"='N' ;";
                }
                else
                {
                    choose.lstrquery = "SELECT  \"CardName\" as \"Name\",\"CardCode\" as \"code\",\"DebPayAcct\"  as \"CtrlCode\" FROM " + ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBNames").Cells.Item(pVal.Row).Specific).Value.ToString() + ".OCRD  where \"validFor\" ='Y' ;";
                }
                Currentrow = pVal.Row;
                choose.Show();
                BubbleEvent = false;
            }

            if (oForm.Mode != BoFormMode.fm_ADD_MODE)
            {
                BubbleEvent = false;
            }
            bool allow = true;
            switch (pVal.CharPressed)
            {
                case 9:
                case 36:
                    allow = false;
                    break;
            }
            if (pVal.ColUID == "U_ODBCredit" && allow)
            {
                if (Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBDebit").Cells.Item(pVal.Row).Specific).Value) !=0)
                {
                    BubbleEvent = false;
                }
            }

            if (pVal.ColUID == "U_ODBDebit" && allow)
            {
                if (Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBCredit").Cells.Item(pVal.Row).Specific).Value) != 0)
                {
                    BubbleEvent = false;
                }
            }


        }

        private void Choose_Retval(SAPbouiCOM.DataTable sender, int Row)
        {
            if (Row >= 0)
            {
                ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBAccNo").Cells.Item(Currentrow).Specific).Value = sender.GetValue("code", Row).ToString();
                ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBAccName").Cells.Item(Currentrow).Specific).Value = sender.GetValue("Name", Row).ToString();
                ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBControlAcc").Cells.Item(Currentrow).Specific).Value = sender.GetValue("CtrlCode", Row).ToString();

            }

        }

        private void Form_DataAddAfter(ref BusinessObjectInfo pVal)
        {

            if (pVal.ActionSuccess)
            {


                Dictionary<string, List<Dictionary<string, object>>> companies = new Dictionary<string, List<Dictionary<string, object>>>();

                for (int i = 0; i < Matrix0.RowCount; i++)
                {
                    string companyName = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBNames").Cells.Item(i + 1).Specific).Value.ToString();
                    Dictionary<string, object> companyData = new Dictionary<string, object>
                    {
                        {"shortname",((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBAccNo").Cells.Item(i+1).Specific).Value.ToString() },
                        {"credit", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBCredit").Cells.Item(i+1).Specific).Value.ToString()},
                        {"debit", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBDebit").Cells.Item(i+1).Specific).Value.ToString()},
                        {"BPLid", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBBranches").Cells.Item(i+1).Specific).Value.ToString()},
                        {"cost1", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_Dim1").Cells.Item(i+1).Specific).Value.ToString()},
                        {"cost2", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_Dim2").Cells.Item(i+1).Specific).Value.ToString()},
                        {"cost3", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_Dim3").Cells.Item(i+1).Specific).Value.ToString()},
                        {"cost4", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_Dim4").Cells.Item(i+1).Specific).Value.ToString()},
                        {"cost5", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_Dim5").Cells.Item(i+1).Specific).Value.ToString()},

                    };

                    int position = new List<string>(companies.Keys).IndexOf(companyName);
                    if (position == -1)
                    {
                        companies.Add(companyName, new List<Dictionary<string, object>> { companyData });
                    }
                    else
                    {
                        companies[companyName].Add(companyData);
                    }
                }

                foreach (var companyName in companies.Keys)
                {
                    Transaction.anotherCompany(companyName, out objAnothercompany, out BPCodeCustomer, out BPCodeVendor);
                    PostVoucher(companies[companyName]);
                }

            }
        }

        public bool saveJourUDF(int DocEntry, string UDFColumn, string UDFValue, SAPbobsCOM.BoObjectTypes boObjectTypes)
        {

            SAPbobsCOM.JournalEntries obj = null;
            obj = (SAPbobsCOM.JournalEntries)clsModule.objaddon.objcompany.GetBusinessObject(boObjectTypes);
            if (obj.GetByKey(DocEntry))
            {
                obj.UserFields.Fields.Item(UDFColumn).Value = UDFValue;
                int ret = obj.Update();
                if (ret != 0)
                {
                    int error = 0;
                    string msg = "";
                    clsModule.objaddon.objcompany.GetLastError(out error, out msg);
                }
            }
            return true;
        }

        private void Matrix0_GotFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (pVal.ColUID == "U_ODBNames" && pVal.Row > 1)
            {
                string previous = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBNames").Cells.Item(pVal.Row - 1).Specific).Value.ToString();
                ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBNames").Cells.Item(pVal.Row).Specific).Value = previous;

            }

            if (pVal.ColUID == "U_ODBBranches" && pVal.Row > 0)
            {
                if (pVal.Row > 1)
                {
                    string previous = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBBranches").Cells.Item(pVal.Row - 1).Specific).Value.ToString();
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBBranches").Cells.Item(pVal.Row).Specific).Value = previous;
                }
            }

            if (pVal.ColUID == "U_ODBBranchName" && pVal.Row > 0)
            {
                string previous = clsModule.objaddon.objglobalmethods.getSingleValue("SELECT \"BPLName\" FROM " + ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBNames").Cells.Item(pVal.Row).Specific).Value.ToString() + ".OBPL WHERE \"BPLId\" ='" + ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBBranches").Cells.Item(pVal.Row).Specific).Value + "'");
                ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBBranchName").Cells.Item(pVal.Row).Specific).Value = previous;
            }

            if (pVal.ColUID == "U_ODBDebit" && pVal.Row > 0)
            {

                decimal Credit = Enumerable.Range(1, Matrix0.RowCount).Sum(row =>
                          Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBDebit").Cells.Item(row).Specific).Value));

                decimal debit = Enumerable.Range(1, Matrix0.RowCount).Sum(row =>
                       Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBCredit").Cells.Item(row).Specific).Value));
                if ((debit - Credit) > 0 && Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBCredit").Cells.Item(pVal.Row).Specific).Value) == 0)
                {
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBDebit").Cells.Item(pVal.Row).Specific).Value = (debit - Credit).ToString();
                }
            }
            if (pVal.ColUID == "U_ODBCredit" && pVal.Row > 0)
            {

                decimal Credit = Enumerable.Range(1, Matrix0.RowCount).Sum(row =>
                          Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBDebit").Cells.Item(row).Specific).Value));

                decimal debit = Enumerable.Range(1, Matrix0.RowCount).Sum(row =>
                       Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBCredit").Cells.Item(row).Specific).Value));

                if ((Credit - debit) > 0 && Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBDebit").Cells.Item(pVal.Row).Specific).Value) == 0)
                {
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBCredit").Cells.Item(pVal.Row).Specific).Value = (Credit - debit).ToString();
                }

            }

            try
            {
                clsModule.objaddon.objapplication.Menus.Item("773").Enabled = true;
                clsModule.objaddon.objapplication.Menus.Item("773").Enabled = oForm.Mode == BoFormMode.fm_ADD_MODE;


               
                if (pVal.ColUID == "U_ODBCredit")
                {
                    if (Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBDebit").Cells.Item(pVal.Row).Specific).Value) != 0)
                    {
                        clsModule.objaddon.objapplication.Menus.Item("773").Enabled =false;
                    }
                }

                if (pVal.ColUID == "U_ODBDebit" )
                {
                    if (Convert.ToDecimal(((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBCredit").Cells.Item(pVal.Row).Specific).Value) != 0)
                    {
                        clsModule.objaddon.objapplication.Menus.Item("773").Enabled = false;
                    }
                }

            }
            catch (Exception ex)
            {

               
            }

           

        }



        private bool PostVoucher(List<Dictionary<string, object>> companies)
        {
            SAPbobsCOM.JournalEntries oJV;
            oJV = (SAPbobsCOM.JournalEntries)objAnothercompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);


            for (int i = 0; i < companies.Count; i++)
            {
                oJV.Lines.ShortName = companies[i]["shortname"].ToString();
                oJV.Lines.Credit = clsModule.objaddon.objglobalmethods.Cton(companies[i]["credit"]);
                oJV.Lines.Debit = clsModule.objaddon.objglobalmethods.Cton(companies[i]["debit"]);
                oJV.Lines.BPLID = clsModule.objaddon.objglobalmethods.Ctoint(companies[i]["BPLid"]);
                oJV.Lines.CostingCode = companies[i]["cost1"].ToString();
                oJV.Lines.CostingCode2 = companies[i]["cost2"].ToString();
                oJV.Lines.CostingCode3 = companies[i]["cost3"].ToString();
                oJV.Lines.CostingCode4 = companies[i]["cost4"].ToString();
                oJV.Lines.CostingCode5 = companies[i]["cost5"].ToString();
                oJV.Lines.Add();
            }
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
                saveJourUDF(Convert.ToInt32(oForm.DataSources.DBDataSources.Item("OJDT").GetValue("TransId", 0)), "U_JouEnt", udfValue, BoObjectTypes.oJournalEntries);
            }

            return true;
        }



        private void Form_ActivateAfter(SBOItemEventArg pVal)
        {
            if (clsModule.objaddon.objglobalmethods.cancel)
            {
                oForm.Freeze(true);
                for (int i = 0; i < Matrix0.RowCount; i++)
                {
                    (((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBDebit").Cells.Item(i + 1).Specific).Value, ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBCredit").Cells.Item(i + 1).Specific).Value) =
                        (((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBCredit").Cells.Item(i + 1).Specific).Value.ToString(),
                        ((SAPbouiCOM.EditText)Matrix0.Columns.Item("U_ODBDebit").Cells.Item(i + 1).Specific).Value.ToString());
                }
                clsModule.objaddon.objglobalmethods.cancel = false;
                oForm.Freeze(false);
            }
        }

        private void Matrix0_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (oForm.Mode != BoFormMode.fm_ADD_MODE)
            {
                BubbleEvent = false;
            }



        }
    }

}
