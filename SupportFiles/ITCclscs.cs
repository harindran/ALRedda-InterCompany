using General.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static General.Common.Module;
namespace ALRedda.Business_Objects.SupportFiles
{
    public class ITCclscs
    {
        private readonly SAPbouiCOM.Form oForm;
        public ITCclscs(SAPbouiCOM.Form Form)
        {
            oForm = Form;
        }
        public void ITCStart()
        {

            SAPbouiCOM.Matrix Matrix0,Matrix1;
            Matrix0 = ((SAPbouiCOM.Matrix)(oForm.Items.Item("MJE").Specific));
            Matrix1 = ((SAPbouiCOM.Matrix)(oForm.Items.Item("Mvendor").Specific));
            SAPbouiCOM.EditText EditText0 = (SAPbouiCOM.EditText)(oForm.Items.Item("DocNum").Specific);
            SAPbouiCOM.EditText EditText1 = (SAPbouiCOM.EditText)(oForm.Items.Item("DocDt").Specific);
            SAPbouiCOM.EditText EditText2 = (SAPbouiCOM.EditText)(oForm.Items.Item("Remark").Specific);
          //  oForm.Settings.Enabled = true;
           
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
            {
                Matrix0.AddRow();
             


                Sapf.EditMatrixcol(Matrix0, oForm, "OffComp", "", "", "", visible: false);
                Sapf.EditMatrixcol(Matrix0, oForm, "OffLed", "", "", "", visible: false);


                EditText1.String = DateTime.Today.ToString("yyyyMMdd");


                string lstr = "";
                string series = "";

                lstr = " SELECT \"Series\"  FROM NNM1 n WHERE \"Indicator\" " +
                    " IN (SELECT \"Indicator\" FROM OFPR WHERE '" + DateTime.Today.ToString("yyyyMMdd") + "' BETWEEN \"F_RefDate\" AND \"T_RefDate\") AND \"ObjectCode\" = 'ATPL_OITC';";

                series = clsModule.objaddon.objglobalmethods.getSingleValue(lstr);

                EditText0.String = clsModule.objaddon.objglobalmethods.GetDocNum("ATPL_OITC", clsModule.objaddon.objglobalmethods.Ctoint(series));

                EditText2.Item.Click();

                EditText0.Item.Enabled = false;
                Matrix0.Columns.Item("#").Editable = false;

                VendorMatrix(Matrix1);
                Matrix1.AddRow();
            }
            else
            {
                Matrix0.Item.Enabled = false;
                Matrix1.Item.Enabled = false;
            }
        }

        public enum colVendor
        {
            U_Comp = 1,
            U_GLCode,
            U_GLName,
            U_GLAcc,
            U_Debit,
            U_Credit,
            U_TaxCode,
            U_TaxName,
            U_TaxRate,
            U_TaxAmt,
            U_Cost1,
            U_Cost2,
            U_Cost3,
            U_Cost4,
            U_Cost5,
            U_OffComp,
            U_OffLed,
            U_OffLedName,
            U_OffDebit,
            U_OffCredit,
            U_OffCost1,
            U_OffCost2,
            U_OffCost3,
            U_OffCost4,
            U_OffCost5,

        }



        private void VendorMatrix(SAPbouiCOM.Matrix Matrix0)
        {
            for (int delrow = 0; delrow < Matrix0.RowCount; delrow++)
            {
                Matrix0.DeleteRow(1);
            }

            for (int i = 1; i < Matrix0.Columns.Count; i++)
            {
                Matrix0.Columns.Remove(1);

            }

            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_Comp).ToString(), "Company", "@ITC2", colVendor.U_Comp.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_GLCode).ToString(), "GL Code", "@ITC2", colVendor.U_GLCode.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_GLName).ToString(), "GL Name", "@ITC2", colVendor.U_GLName.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_GLAcc).ToString(), "GL Account", "@ITC2", colVendor.U_GLAcc.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_Debit).ToString(), "Debit", "@ITC2", colVendor.U_Debit.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_Credit).ToString(), "Credit", "@ITC2", colVendor.U_Credit.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_TaxCode).ToString(), "TaxCode", "@ITC2", colVendor.U_TaxCode.ToString(), 75,Types:SAPbouiCOM.BoFormItemTypes.it_EDIT,cfl:oForm.ChooseFromLists.Item("CFL_Tax"));
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_TaxName).ToString(), "TaxName", "@ITC2", colVendor.U_TaxName.ToString(), 75,visible:false);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_TaxRate).ToString(), "TaxRate", "@ITC2", colVendor.U_TaxRate.ToString(), 75,visible:false);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_TaxAmt).ToString(), "TaxAmt", "@ITC2", colVendor.U_TaxAmt.ToString(), 75,Edit:false);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_Cost1).ToString(), "Cost Center1", "@ITC2", colVendor.U_Cost1.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_Cost2).ToString(), "Cost Center2", "@ITC2", colVendor.U_Cost2.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_Cost3).ToString(), "Cost Center3", "@ITC2", colVendor.U_Cost3.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_Cost4).ToString(), "Cost Center4", "@ITC2", colVendor.U_Cost4.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_Cost5).ToString(), "Cost Center5", "@ITC2", colVendor.U_Cost5.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_OffComp).ToString(), "OffSet Company", "@ITC2", colVendor.U_OffComp.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_OffLed).ToString(), "OffSet Ledger Code", "@ITC2", colVendor.U_OffLed.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_OffLedName).ToString(), "OffSet Ledger Name", "@ITC2", colVendor.U_OffLedName.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_OffDebit).ToString(), "OffSet Debit", "@ITC2", colVendor.U_OffDebit.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_OffCredit).ToString(), "OffSet Credit", "@ITC2", colVendor.U_OffCredit.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_OffCost1).ToString(), "OffSet Cost Center1", "@ITC2", colVendor.U_OffCost1.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_OffCost2).ToString(), "OffSet Cost Center2", "@ITC2", colVendor.U_OffCost2.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_OffCost3).ToString(), "OffSet Cost Center3", "@ITC2", colVendor.U_OffCost3.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_OffCost4).ToString(), "OffSet Cost Center4", "@ITC2", colVendor.U_OffCost4.ToString(), 75);
            Sapf.AddMatrixcol(Matrix0, oForm, "MVen" + ((int)colVendor.U_OffCost5).ToString(), "OffSet Cost Center5", "@ITC2", colVendor.U_OffCost5.ToString(), 75);

            Matrix0.Item.Visible = false;
            Sapf.EditMatrixcol(Matrix0, oForm, "#", "#", "","", Edit: false);

        }

    }
}
