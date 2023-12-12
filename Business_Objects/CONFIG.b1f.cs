using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using General.Common;
using SAPbobsCOM;
using SAPbouiCOM.Framework;

namespace ALRedda.Business_Objects
{
    [FormAttribute("CONFIG", "Business_Objects/CONFIG.b1f")]
    public class CONFIG : UserFormBase
    {
        public SAPbouiCOM.Form oForm;
        public CONFIG()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_0").Specific));
            this.Matrix0.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix0_ValidateAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new LoadAfterHandler(this.Form_LoadAfter);
            
        }

        private SAPbouiCOM.Matrix Matrix0;

        private void OnCustomInitialize()
        {
            Loaddata();
            Matrix0.AddRow();
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
    }

        private void Loaddata()
        {

            string strSQL = "";                        
            try
            {

                strSQL = "Select \"U_DBName\",\"U_DBUser\",\"U_DBPass\",\"U_sysUser\",\"U_sysPass\",\"U_BPCode\",\"U_BPCode1\" ";
                strSQL += @" from ""@CONFIG"" T0 join ""@CONFIG1"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='01'";                
                SAPbobsCOM.Recordset rc = clsModule.objaddon.objglobalmethods.GetmultipleRS(strSQL);
                Matrix0.Clear();
                for (int i = 0; i < rc.RecordCount; i++)
                {
                    Matrix0.AddRow();
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("DBName").Cells.Item(i+1).Specific).Value = rc.Fields.Item("U_DBName").Value.ToString() ;
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("DBUser").Cells.Item(i+1).Specific).Value = rc.Fields.Item("U_DBUser").Value.ToString() ;
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("DBPass").Cells.Item(i+1).Specific).Value = rc.Fields.Item("U_DBPass").Value.ToString() ;
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("User").Cells.Item(i+1).Specific).Value = rc.Fields.Item("U_sysUser").Value.ToString() ;
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Pass").Cells.Item(i+1).Specific).Value = rc.Fields.Item("U_sysPass").Value.ToString() ;
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("BPCode").Cells.Item(i+1).Specific).Value = rc.Fields.Item("U_BPCode").Value.ToString() ;
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("BPCode1").Cells.Item(i+1).Specific).Value = rc.Fields.Item("U_BPCode1").Value.ToString() ;
                    }                    
                }            

            catch (Exception ex)
            {
              
            }


        }

        private bool Save()
        {
            try
            {
                bool Flag = false;
                //string live;
                GeneralService oGeneralService;
                GeneralData oGeneralData;
                GeneralDataParams oGeneralParams;
                GeneralDataCollection oGeneralDataCollection;
                GeneralData oChild;

                oGeneralService = clsModule.objaddon.objcompany.GetCompanyService().GetGeneralService("ATPL_CONFIG");
                oGeneralData = (GeneralData)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);
                oGeneralParams = (GeneralDataParams)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralDataCollection = oGeneralData.Child("CONFIG1");
                try
                {
                    oGeneralParams.SetProperty("Code", "01");
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                    Flag = true;
                }
                catch (Exception ex)
                {
                    Flag = false;
                }
                oGeneralData.SetProperty("Code", "01");
                oGeneralData.SetProperty("Name", "01");             
                oChild = oGeneralDataCollection.Add();
                int rowcount = 0;
                for (int i = 1; i <= Matrix0.VisualRowCount; i++)
                {
                    if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("DBName").Cells.Item(i).Specific).String != "")
                    {
                        oGeneralData.Child("CONFIG1").Item(rowcount).SetProperty("U_DBName", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("DBName").Cells.Item(i).Specific).String);
                        oGeneralData.Child("CONFIG1").Item(rowcount).SetProperty("U_DBUser", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("DBUser").Cells.Item(i).Specific).String);
                        oGeneralData.Child("CONFIG1").Item(rowcount).SetProperty("U_DBPass", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("DBPass").Cells.Item(i).Specific).String);
                        oGeneralData.Child("CONFIG1").Item(rowcount).SetProperty("U_sysUser", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("User").Cells.Item(i).Specific).String);
                        oGeneralData.Child("CONFIG1").Item(rowcount).SetProperty("U_sysPass", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Pass").Cells.Item(i).Specific).String);
                        oGeneralData.Child("CONFIG1").Item(rowcount).SetProperty("U_BPCode", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("BPCode").Cells.Item(i).Specific).String);
                        oGeneralData.Child("CONFIG1").Item(rowcount).SetProperty("U_BPCode1", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("BPCode1").Cells.Item(i).Specific).String);
                        rowcount++;
                    }
                }
                
                if (Flag == true)
                {
                    oGeneralService.Update(oGeneralData);
                    return true;
                }
                else
                {
                    oGeneralParams = oGeneralService.Add(oGeneralData);
                    return true;
                }
            }

            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm = clsModule.objaddon.objapplication.Forms.GetForm("CONFIG", pVal.FormTypeCount);

        }

        private SAPbouiCOM.Button Button0;

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;           
        }

        private void Button0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            Save();
        }

        private void Matrix0_ValidateAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            
                try
                {                   
                    switch (pVal.ColUID)
                    {
                        case "Pass":
                            clsModule.objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "Pass", "#");
                            break;
                    }

                }
                catch (Exception)
                {
                    throw;
                }
            

        }
    }
}
