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

            SAPbouiCOM.Matrix Matrix0;
            Matrix0 = ((SAPbouiCOM.Matrix)(oForm.Items.Item("Item_0").Specific));
            SAPbouiCOM.EditText EditText0 = (SAPbouiCOM.EditText)(oForm.Items.Item("DocNum").Specific);
            SAPbouiCOM.EditText EditText1 = (SAPbouiCOM.EditText)(oForm.Items.Item("DocDt").Specific);
            SAPbouiCOM.EditText EditText2 = (SAPbouiCOM.EditText)(oForm.Items.Item("Remark").Specific);

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
            }
            else
            {
                Matrix0.Item.Enabled = false;
            }
        }


    }
}
