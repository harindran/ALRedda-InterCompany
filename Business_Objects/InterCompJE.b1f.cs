using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ALRedda.Business_Objects.SupportFiles;
using General.Common;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using static General.Common.Module;
using static ALRedda.Business_Objects.SupportFiles.ITCclscs;

namespace ALRedda.Business_Objects
{
    [FormAttribute("ICT", "Business_Objects/InterCompJE.b1f")]
    class InterCompJE : UserFormBase
    {
        public SAPbouiCOM.Form oForm;
        private int Currentrow = 0;
        private SAPbobsCOM.Company objAnothercompany;
        private string BPCodeCustomer = "";
        private string BPCodeVendor = "";
        private string DocEntry = "";
        SAPbouiCOM.DBDataSource odbHeader,ODbvender;
        bool offsetcomp = false;
       
        public InterCompJE()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("MJE").Specific));
            this.Matrix0.GotFocusAfter += new SAPbouiCOM._IMatrixEvents_GotFocusAfterEventHandler(this.Matrix0_GotFocusAfter);
            this.Matrix0.KeyDownBefore += new SAPbouiCOM._IMatrixEvents_KeyDownBeforeEventHandler(this.Matrix0_KeyDownBefore);
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_2").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button0.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.Button0.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button0_ClickAfter);
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("LDocNo").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_4").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("DocNum").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("DocDt").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_7").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("Remark").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("DocEntry").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("Item_3").Specific));
            this.Button2.ClickAfter += new SAPbouiCOM._IButtonEvents_ClickAfterEventHandler(this.Button2_ClickAfter);
            this.Button2.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button2_ClickBefore);
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("Mvendor").Specific));
            this.Matrix1.ValidateAfter += new SAPbouiCOM._IMatrixEvents_ValidateAfterEventHandler(this.Matrix1_ValidateAfter);
            this.Matrix1.LostFocusAfter += new SAPbouiCOM._IMatrixEvents_LostFocusAfterEventHandler(this.Matrix1_LostFocusAfter);
            this.Matrix1.ChooseFromListAfter += new SAPbouiCOM._IMatrixEvents_ChooseFromListAfterEventHandler(this.Matrix1_ChooseFromListAfter);
            this.Matrix1.GotFocusAfter += new SAPbouiCOM._IMatrixEvents_GotFocusAfterEventHandler(this.Matrix1_GotFocusAfter);
            this.Matrix1.KeyDownBefore += new SAPbouiCOM._IMatrixEvents_KeyDownBeforeEventHandler(this.Matrix1_KeyDownBefore);
            this.EditText4 = ((SAPbouiCOM.EditText)(this.GetItem("Tvendor").Specific));
            this.EditText4.LostFocusAfter += new SAPbouiCOM._IEditTextEvents_LostFocusAfterEventHandler(this.EditText4_LostFocusAfter);
            this.EditText4.ChooseFromListAfter += new SAPbouiCOM._IEditTextEvents_ChooseFromListAfterEventHandler(this.EditText4_ChooseFromListAfter);
            this.EditText4.ChooseFromListBefore += new SAPbouiCOM._IEditTextEvents_ChooseFromListBeforeEventHandler(this.EditText4_ChooseFromListBefore);
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_8").Specific));
            this.LinkedButton0 = ((SAPbouiCOM.LinkedButton)(this.GetItem("Item_9").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_10").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("TCardName").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.DataAddAfter += new SAPbouiCOM.Framework.FormBase.DataAddAfterHandler(this.Form_DataAddAfter);
            this.DataLoadAfter += new DataLoadAfterHandler(this.Form_DataLoadAfter);

        }

        private SAPbouiCOM.Matrix Matrix0;



        private void OnCustomInitialize()
        {
            // startInit();
            ITCclscs cNCReq = new ITCclscs(oForm);
            cNCReq.ITCStart();

            odbHeader = oForm.DataSources.DBDataSources.Item("@OITC");
            ODbvender = oForm.DataSources.DBDataSources.Item("@ITC2");

        }

        private void startInit()
        {
            Matrix0.AddRow();
            

            EditText1.String = DateTime.Today.ToString("yyyyMMdd");

            string lstr = "";
            string series = "";

            lstr = " SELECT \"Series\"  FROM NNM1 n WHERE \"Indicator\" " +
                " IN (SELECT \"Indicator\" FROM OFPR WHERE '" + DateTime.Today.ToString("yyyyMMdd") + "' BETWEEN \"F_RefDate\" AND \"T_RefDate\") AND \"ObjectCode\" = 'ATPL_OITC';";

            series = clsModule.objaddon.objglobalmethods.getSingleValue(lstr);

            EditText0.String = clsModule.objaddon.objglobalmethods.GetDocNum("ATPL_OITC", stf.Ctoint(series));

            EditText2.Item.Click();

            EditText0.Item.Enabled = false;
            Matrix0.Columns.Item("#").Editable = false;
        }

        private SAPbouiCOM.Folder Folder0;

        private void Choose_Retval(SAPbouiCOM.DataTable sender, int Row)
        {
            if (string.IsNullOrEmpty(EditText4.Value))
            {
                if (Row >= 0)
                {
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("GLCode").Cells.Item(Currentrow).Specific).Value = sender.GetValue("code", Row).ToString();
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("GLName").Cells.Item(Currentrow).Specific).Value = sender.GetValue("Name", Row).ToString();
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("GLAcc").Cells.Item(Currentrow).Specific).Value = sender.GetValue("CtrlCode", Row).ToString();
                }
            }
            else
            {
                if (Row >= 0)
                {
                    if (offsetcomp)
                    {
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_OffLed).ToString()).Cells.Item(Currentrow).Specific).Value = sender.GetValue("code", Row).ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_OffLedName).ToString()).Cells.Item(Currentrow).Specific).Value = sender.GetValue("Name", Row).ToString();

                    }
                    else
                    {
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_GLCode).ToString()).Cells.Item(Currentrow).Specific).Value = sender.GetValue("code", Row).ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_GLName).ToString()).Cells.Item(Currentrow).Specific).Value = sender.GetValue("Name", Row).ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_GLAcc).ToString()).Cells.Item(Currentrow).Specific).Value = sender.GetValue("CtrlCode", Row).ToString();
                    }
                }
            }
        }


        private void Matrix0_KeyDownBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {

            BubbleEvent = true;

            if (pVal.ColUID == "GLCode" && pVal.CharPressed == 9)
            {

                if (!string.IsNullOrEmpty(((SAPbouiCOM.EditText)Matrix0.Columns.Item("GLCode").Cells.Item(pVal.Row).Specific).Value)) return;
                if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)Matrix0.Columns.Item("DBComp").Cells.Item(pVal.Row).Specific).Value)) return;
                choose choose = new choose();
                choose.Retval += Choose_Retval;

                if (pVal.Modifiers != SAPbouiCOM.BoModifiersEnum.mt_CTRL)
                {
                    choose.lstrquery = "SELECT  \"AcctName\" as \"Name\" ,\"AcctCode\" as \"code\", \"AcctCode\" as \"CtrlCode\" FROM " + ((SAPbouiCOM.EditText)Matrix0.Columns.Item("DBComp").Cells.Item(pVal.Row).Specific).Value.ToString() + ".OACT  where \"LocManTran\" ='N' and  \"Postable\" ='Y'  AND \"FrozenFor\"='N'  order by \"Name\";";
                }
                else
                {
                    choose.lstrquery = "SELECT  \"CardName\" as \"Name\",\"CardCode\" as \"code\",\"DebPayAcct\"  as \"CtrlCode\" FROM " + ((SAPbouiCOM.EditText)Matrix0.Columns.Item("DBComp").Cells.Item(pVal.Row).Specific).Value.ToString() + ".OCRD  where \"validFor\" ='Y' order by \"Name\" ;";
                }
                Currentrow = pVal.Row;
                choose.Show();
                Matrix0.AddRow();
                BubbleEvent = false;
            }

            if (pVal.ColUID == "OffComp" && pVal.CharPressed == 9)
            {
                if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)Matrix0.Columns.Item("OffComp").Cells.Item(pVal.Row).Specific).Value)) return;
                if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)Matrix0.Columns.Item("DBComp").Cells.Item(pVal.Row).Specific).Value)) return;

                string DB2 = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OffComp").Cells.Item(pVal.Row).Specific).Value;
                string DB1 = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("DBComp").Cells.Item(pVal.Row).Specific).Value;
                string lstquery = "SELECT \"U_DBOffset\"  FROM \"@CONFIG2\" c WHERE \"U_DBName1\" ='" + DB1 + "' AND \"U_DBName2\" ='" + DB2 + "' ;";
                ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OffLed").Cells.Item(Currentrow).Specific).Value = clsModule.objaddon.objglobalmethods.getSingleValue(lstquery);

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
            if (pVal.ColUID == "Credit" && allow)
            {
                if (clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Debit").Cells.Item(pVal.Row).Specific).Value) != 0)
                {
                    BubbleEvent = false;
                }
            }

            if (pVal.ColUID == "Debit" && allow)
            {
                if (clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Credit").Cells.Item(pVal.Row).Specific).Value) != 0)
                {
                    BubbleEvent = false;
                }
            }

        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            oForm = clsModule.objaddon.objapplication.Forms.GetForm("ICT", pVal.FormTypeCount);

        }

        private void Matrix0_GotFocusAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (pVal.ColUID == "GLName" && pVal.Row > 1)
            {
                string previous = ((SAPbouiCOM.EditText)Matrix0.Columns.Item("GLName").Cells.Item(pVal.Row - 1).Specific).Value.ToString();
                ((SAPbouiCOM.EditText)Matrix0.Columns.Item("GLName").Cells.Item(pVal.Row).Specific).Value = previous;

            }


            if (pVal.ColUID == "Debit" && pVal.Row > 0)
            {

                decimal Credit = Enumerable.Range(1, Matrix0.RowCount).Sum(row =>
                          clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Debit").Cells.Item(row).Specific).Value));

                decimal debit = Enumerable.Range(1, Matrix0.RowCount).Sum(row =>
                      clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Credit").Cells.Item(row).Specific).Value));
                if ((debit - Credit) > 0 && clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Credit").Cells.Item(pVal.Row).Specific).Value) == 0)
                {
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Debit").Cells.Item(pVal.Row).Specific).Value = (debit - Credit).ToString();
                }
            }
            if (pVal.ColUID == "Credit" && pVal.Row > 0)
            {

                decimal Credit = Enumerable.Range(1, Matrix0.RowCount).Sum(row =>
                         clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Debit").Cells.Item(row).Specific).Value));

                decimal debit = Enumerable.Range(1, Matrix0.RowCount).Sum(row =>
                       clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Credit").Cells.Item(row).Specific).Value));

                if ((Credit - debit) > 0 && clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Debit").Cells.Item(pVal.Row).Specific).Value) == 0)
                {
                    ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Credit").Cells.Item(pVal.Row).Specific).Value = (Credit - debit).ToString();
                }

            }

            try
            {
                clsModule.objaddon.objapplication.Menus.Item("773").Enabled = true;
                clsModule.objaddon.objapplication.Menus.Item("773").Enabled = oForm.Mode == BoFormMode.fm_ADD_MODE;



                if (pVal.ColUID == "Credit")
                {
                    if (clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Debit").Cells.Item(pVal.Row).Specific).Value) != 0)
                    {
                        clsModule.objaddon.objapplication.Menus.Item("773").Enabled = false;
                    }
                }

                if (pVal.ColUID == "Debit")
                {
                    if (clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Credit").Cells.Item(pVal.Row).Specific).Value) != 0)
                    {
                        clsModule.objaddon.objapplication.Menus.Item("773").Enabled = false;
                    }
                }

            }
            catch (Exception ex)
            {


            }
       
        }

        private bool PostVoucher(List<Dictionary<string, object>> company)
        {
            SAPbobsCOM.JournalEntries oJV;
            oJV = (SAPbobsCOM.JournalEntries)objAnothercompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);


            var companies = (company
         .GroupBy(data => new
         {
             ShortName = data["shortname"],
             cost1 = data["cost1"],
             cost2 = data["cost2"],
             cost3 = data["cost3"],
             cost4 = data["cost4"],
             cost5 = data["cost5"],
             ReferenceDate = data["refDate"],
             Remark = data["Remark"],
         })

         .Select(value =>
         {
             return new
             {
                 GroupKey = value.Key.ShortName,
                 CreditSum = value.Sum(item => stf.CtoD(item["credit"])),
                 DebitSum = value.Sum(item => stf.CtoD(item["debit"])),
                 cost1 = stf.ObjtoStr(value.Key.cost1),
                 cost2 = stf.ObjtoStr(value.Key.cost2),
                 cost3 = stf.ObjtoStr(value.Key.cost3),
                 cost4 = stf.ObjtoStr(value.Key.cost4),
                 cost5 = stf.ObjtoStr(value.Key.cost5),
                 ReferenceDate = stf.GetDate(value.Key.ReferenceDate),
                 Remark = stf.ObjtoStr(value.Key.Remark)
             };
         })).ToList();

            for (int i = 0; i < companies.Count; i++)
            {
                decimal credit = 0;
                decimal Debit = 0;

                decimal totval = clsModule.objaddon.objglobalmethods.CtoD(companies[i].CreditSum) - clsModule.objaddon.objglobalmethods.CtoD(companies[i].DebitSum);

                if (totval > 0)
                    credit = totval;
                else
                    Debit = Math.Abs(totval);


                oJV.ReferenceDate = stf.GetDate(companies[i].ReferenceDate.ToString());
                oJV.Memo = companies[i].Remark.ToString();

                oJV.Lines.ShortName = companies[i].GroupKey.ToString();
                oJV.Lines.Credit = clsModule.objaddon.objglobalmethods.Cton(credit);
                oJV.Lines.Debit = clsModule.objaddon.objglobalmethods.Cton(Debit); 
                oJV.Lines.BPLID = 1;
                oJV.Lines.CostingCode = companies[i].cost1.ToString();
                oJV.Lines.CostingCode2 = companies[i].cost2.ToString();
                oJV.Lines.CostingCode3 = companies[i].cost3.ToString();
                oJV.Lines.CostingCode4 = companies[i].cost4.ToString();
                oJV.Lines.CostingCode5 = companies[i].cost5.ToString();

                oJV.Lines.Add();
            }
            int iErrCode=0;
           iErrCode = oJV.Add();
            string strerr = "";
            if (iErrCode != 0)
            {
                objAnothercompany.GetLastError(out iErrCode, out strerr);
                clsModule.objaddon.objapplication.MessageBox(strerr);
            }
            else
            {
                string udfValue = objAnothercompany.GetNewObjectKey();
            }

            return true;
        }


        private Button Button0;
        private Button Button1;


        private bool PostotherDB()
        {
            string Crdcode = oForm.DataSources.DBDataSources.Item("@OITC").GetValue("U_CardCode", 0);
            if (string.IsNullOrEmpty(Crdcode))
            {
                InterCompanyTransaction();
            }
            else
            {
                SupplierTransaction();
            }
            return true;
        }

        private void InterCompanyTransaction()
        {
            string lstrquery = "";

            lstrquery += "  SELECT \"U_DBComp\" ,\"U_GLCode\" ,\"U_GLName\" ,\"U_GLAcc\" ,\"U_Debit\" ,\"U_Credit\" ,\"U_OffComp\" ,\"U_OffLed\", ";
            lstrquery += "  \"U_Dim1\" ,\"U_Dim2\" ,\"U_Dim3\" ,\"U_Dim4\" ,\"U_Dim5\",\"U_DocDate\",\"U_Remarks\"  FROM \"@OITC\" t1 ";
            lstrquery += "  LEFT JOIN \"@ITC1\" t2  ON t1.\"DocEntry\" =t2.\"DocEntry\" where t1.\"DocEntry\"='" + DocEntry + "' ; ";
            SAPbobsCOM.Recordset rc = clsModule.objaddon.objglobalmethods.GetmultipleRS(lstrquery);
            Dictionary<string, List<Dictionary<string, object>>> companies = new Dictionary<string, List<Dictionary<string, object>>>();

            if (rc.RecordCount > 0)
            {

                for (int i = 0; i < rc.RecordCount; i++)
                {
                    string companyName = rc.Fields.Item("U_DBComp").Value.ToString();
                    string Code;
                    string Credit;
                    string Debit;
                    string Dim1;
                    string Dim2;
                    string Dim3;
                    string Dim4;
                    string Dim5;
                    string refdate;
                    string Remark;

                    string currcredit = "";
                    string currdebit = "";
                    if (stf.CtoD(rc.Fields.Item("U_Credit").Value.ToString()) != 0)
                    {
                        currcredit = rc.Fields.Item("U_Credit").Value.ToString();
                    }
                    else
                    {
                        currdebit = rc.Fields.Item("U_Debit").Value.ToString();
                    }

                    for (int j = 0; j < 2; j++)
                    {
                        string OffsetLed = rc.Fields.Item("U_OffLed").Value.ToString();

                        if (string.IsNullOrEmpty(rc.Fields.Item("U_OffLed").Value.ToString()) && j != 0)
                        {
                            string lstquery = "SELECT \"U_DBOffset\"  FROM \"@CONFIG2\" c WHERE \"U_DBName1\" ='" + rc.Fields.Item("U_DBComp").Value.ToString() + "';";
                            OffsetLed = clsModule.objaddon.objglobalmethods.getSingleValue(lstquery);
                        }

                        Code = j == 0 ? rc.Fields.Item("U_GLCode").Value.ToString() : OffsetLed;
                        Credit = j == 0 ? currcredit : currdebit;
                        Debit = j == 0 ? currdebit : currcredit;
                        Dim1 = rc.Fields.Item("U_Dim1").Value.ToString();
                        Dim2 = rc.Fields.Item("U_Dim2").Value.ToString();
                        Dim3 = rc.Fields.Item("U_Dim3").Value.ToString();
                        Dim4 = rc.Fields.Item("U_Dim4").Value.ToString();
                        Dim5 = rc.Fields.Item("U_Dim5").Value.ToString();
                        refdate = rc.Fields.Item("U_DocDate").Value.ToString();
                        Remark = rc.Fields.Item("U_Remarks").Value.ToString();


                        Dictionary<string, object> companyData = new Dictionary<string, object>
                    {
                        {"shortname",Code },
                        {"credit", Credit},
                        {"debit", Debit},
                        {"cost1", Dim1},
                        {"cost2", Dim2},
                        {"cost3", Dim3},
                        {"cost4", Dim4},
                        {"cost5", Dim5},
                        { "refDate",refdate},
                        { "Remark",Remark }

                    };

                        int position = new List<string>(companies.Keys).IndexOf(companyName);
                        if (position == -1)
                        {
                            companies.Add(companyName, new List<Dictionary<string, object>> { companyData });
                            clsModule.objaddon.objapplication.StatusBar.SetText(companyName + "processing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                        }
                        else
                        {
                            companies[companyName].Add(companyData);
                        }
                    }

                    rc.MoveNext();
                }



            }


            companies.Remove(clsModule.objaddon.objcompany.CompanyDB);

            //Current Company

            lstrquery = "";

            lstrquery += "  SELECT '" + clsModule.objaddon.objcompany.CompanyDB + "' \"U_DBComp\",\"U_DBComp\" as \"MainDB\", ";
            lstrquery += " CASE WHEN \"U_DBComp\" <>  '" + clsModule.objaddon.objcompany.CompanyDB + "'THEN '' ELSE \"U_GLCode\" END AS \"U_GLCode\" ";
            lstrquery += ",\"U_GLName\" ,\"U_GLAcc\" ,\"U_Debit\" ,\"U_Credit\" ,\"U_OffComp\" ,\"U_OffLed\", ";
            lstrquery += "  \"U_Dim1\" ,\"U_Dim2\" ,\"U_Dim3\" ,\"U_Dim4\" ,\"U_Dim5\",\"U_DocDate\",\"U_Remarks\"  FROM \"@OITC\" t1 ";
            lstrquery += "  LEFT JOIN \"@ITC1\" t2  ON t1.\"DocEntry\" =t2.\"DocEntry\" where t1.\"DocEntry\"='" + DocEntry + "' ; ";
            rc = clsModule.objaddon.objglobalmethods.GetmultipleRS(lstrquery);
            if (rc.RecordCount > 0)
            {

                for (int i = 0; i < rc.RecordCount; i++)
                {
                    string companyName = rc.Fields.Item("U_DBComp").Value.ToString();
                    string Code;
                    string Credit;
                    string Debit;
                    string Dim1;
                    string Dim2;
                    string Dim3;
                    string Dim4;
                    string Dim5;
                    string refdate;
                    string Remark;

                    string currcredit = "";
                    string currdebit = "";
                    if (stf.CtoD(rc.Fields.Item("U_Credit").Value.ToString()) != 0)
                    {
                        currcredit = rc.Fields.Item("U_Credit").Value.ToString();
                    }
                    else
                    {
                        currdebit = rc.Fields.Item("U_Debit").Value.ToString();
                    }


                    string OffsetLed = rc.Fields.Item("U_GLCode").Value.ToString();

                    if (string.IsNullOrEmpty(rc.Fields.Item("U_GLCode").Value.ToString()))
                    {
                        string lstquery = "SELECT \"U_DBOffset\"  FROM \"@CONFIG2\" c WHERE \"U_DBName1\" ='" + rc.Fields.Item("MainDB").Value.ToString() + "';";
                        OffsetLed = clsModule.objaddon.objglobalmethods.getSingleValue(lstquery);
                    }

                    Code = OffsetLed;
                    Credit = currcredit;
                    Debit = currdebit;
                    Dim1 = rc.Fields.Item("U_Dim1").Value.ToString();
                    Dim2 = rc.Fields.Item("U_Dim2").Value.ToString();
                    Dim3 = rc.Fields.Item("U_Dim3").Value.ToString();
                    Dim4 = rc.Fields.Item("U_Dim4").Value.ToString();
                    Dim5 = rc.Fields.Item("U_Dim5").Value.ToString();
                    refdate = rc.Fields.Item("U_DocDate").Value.ToString();
                    Remark = rc.Fields.Item("U_Remarks").Value.ToString();


                    Dictionary<string, object> companyData = new Dictionary<string, object>
                    {
                        {"shortname",Code },
                        {"credit", Credit},
                        {"debit", Debit},
                        {"cost1", Dim1},
                        {"cost2", Dim2},
                        {"cost3", Dim3},
                        {"cost4", Dim4},
                        {"cost5", Dim5},
                        { "refDate",refdate},
                        { "Remark",Remark }

                    };

                    int position = new List<string>(companies.Keys).IndexOf(companyName);
                    if (position == -1)
                    {
                        companies.Add(companyName, new List<Dictionary<string, object>> { companyData });
                        clsModule.objaddon.objapplication.StatusBar.SetText(companyName + "processing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    }
                    else
                    {
                        companies[companyName].Add(companyData);
                    }


                    rc.MoveNext();
                }
            }




            foreach (var companyName in companies.Keys)
            {
                if (!string.IsNullOrEmpty(companyName))
                {

                    clsModule.objaddon.objapplication.StatusBar.SetText("Starting Connection " + companyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    Transaction.anotherCompany(companyName, out objAnothercompany, out BPCodeCustomer, out BPCodeVendor);
                    PostVoucher(companies[companyName]);
                }
            }
        }

        private bool SupplierTransaction()
        {
            string lstrquery = "";

            lstrquery += "  SELECT \"U_OffComp\" \"U_DBComp\" ,\"U_OffLed\" \"U_GLCode\" ,\"U_OffDebit\"  \"U_Debit\" ,\"U_OffCredit\" \"U_Credit\", ";
            lstrquery += " \"U_OffCost1\" \"U_Dim1\" ,\"U_OffCost2\" \"U_Dim2\" ,\"U_OffCost3\" \"U_Dim3\" ,\"U_OffCost4\" \"U_Dim4\" ,\"U_OffCost5\" \"U_Dim5\",\"U_DocDate\",\"U_Remarks\"  FROM \"@OITC\" t1 ";
            lstrquery += "  LEFT JOIN \"@ITC2\" t2  ON t1.\"DocEntry\" =t2.\"DocEntry\" where t1.\"DocEntry\"='" + DocEntry + "' ; ";
            SAPbobsCOM.Recordset rc = clsModule.objaddon.objglobalmethods.GetmultipleRS(lstrquery);
            Dictionary<string, List<Dictionary<string, object>>> companies = new Dictionary<string, List<Dictionary<string, object>>>();

            if (rc.RecordCount > 0)
            {

                for (int i = 0; i < rc.RecordCount; i++)
                {
                    string companyName = rc.Fields.Item("U_DBComp").Value.ToString();
                    string Code;
                    string Credit;
                    string Debit;
                    string Dim1;
                    string Dim2;
                    string Dim3;
                    string Dim4;
                    string Dim5;
                    string refdate;
                    string Remark;

                    string currcredit = "";
                    string currdebit = "";
                    if (stf.CtoD(rc.Fields.Item("U_Credit").Value.ToString()) != 0)
                    {
                        currcredit = rc.Fields.Item("U_Credit").Value.ToString();
                    }
                    else
                    {
                        currdebit = rc.Fields.Item("U_Debit").Value.ToString();
                    }

                    for (int j = 0; j < 2; j++)
                    {
                        string OffsetLed = rc.Fields.Item("U_GLCode").Value.ToString(); 

                        if (string.IsNullOrEmpty(rc.Fields.Item("U_GLCode").Value.ToString())&& j != 0)
                        {
                            string lstquery = "SELECT \"U_DBOffset\"  FROM \"@CONFIG2\" c WHERE \"U_DBName1\" ='" + rc.Fields.Item("U_DBComp").Value.ToString() + "';";
                            OffsetLed = clsModule.objaddon.objglobalmethods.getSingleValue(lstquery);
                        }

                        Code = j == 0 ? rc.Fields.Item("U_GLCode").Value.ToString() : OffsetLed;
                        Credit = j == 0 ? currcredit : currdebit;
                        Debit = j == 0 ? currdebit : currcredit;
                        Dim1 = rc.Fields.Item("U_Dim1").Value.ToString();
                        Dim2 = rc.Fields.Item("U_Dim2").Value.ToString();
                        Dim3 = rc.Fields.Item("U_Dim3").Value.ToString();
                        Dim4 = rc.Fields.Item("U_Dim4").Value.ToString();
                        Dim5 = rc.Fields.Item("U_Dim5").Value.ToString();
                        refdate = rc.Fields.Item("U_DocDate").Value.ToString();
                        Remark = rc.Fields.Item("U_Remarks").Value.ToString();


                        Dictionary<string, object> companyData = new Dictionary<string, object>
                    {
                        {"shortname",Code },
                        {"credit", Credit},
                        {"debit", Debit},
                        {"cost1", Dim1},
                        {"cost2", Dim2},
                        {"cost3", Dim3},
                        {"cost4", Dim4},
                        {"cost5", Dim5},
                        { "refDate",refdate},
                        { "Remark",Remark }

                    };

                        int position = new List<string>(companies.Keys).IndexOf(companyName);
                        if (position == -1)
                        {
                            companies.Add(companyName, new List<Dictionary<string, object>> { companyData });
                            clsModule.objaddon.objapplication.StatusBar.SetText(companyName + "processing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                        }
                        else
                        {
                            companies[companyName].Add(companyData);
                        }
                    }

                    rc.MoveNext();
                }



            }


            companies.Remove(clsModule.objaddon.objcompany.CompanyDB);

            //Current Company

            lstrquery = "";

            lstrquery += " SELECT \"U_Comp\" \"U_DBComp\",\"U_GLCode\" ,'CurrDB' AS types,sum((\"U_Debit\" -\"U_OffDebit\" )-(\"U_Credit\" -\"U_OffCredit\")) \"Amount\" ,'' \"MainDB\", ";
            lstrquery += " \"U_Cost1\" \"U_Dim1\" ,\"U_Cost2\" \"U_Dim2\",\"U_Cost3\" \"U_Dim3\",\"U_Cost4\" \"U_Dim4\" ,\"U_Cost5\" \"U_Dim5\",t1.\"U_DocDate\",t1.\"U_Remarks\" FROM \"@ITC2\" i ";
            lstrquery += " left join \"@OITC\" t1 on t1.\"DocEntry\" = i.\"DocEntry\" ";
            lstrquery += "  where i.\"DocEntry\"='" + DocEntry + "'";
            lstrquery += " GROUP BY \"U_Comp\",\"U_GLCode\" , \"U_Cost1\" , \"U_Cost2\" , \"U_Cost3\" , \"U_Cost4\" , \"U_Cost5\" , ";
            lstrquery += " t1.\"U_DocDate\",t1.\"U_Remarks\" ";
            lstrquery += " UNION ALL ";
            lstrquery += " SELECT \"U_Comp\" ,o.\"Account\"  ,'Tax' AS types, ";
            lstrquery += " sum(\"U_TaxAmt\"),'' \"MainDB\", '', '', '', '', '',t1.\"U_DocDate\",t1.\"U_Remarks\" FROM \"@ITC2\" i ";
            lstrquery += " left join \"@OITC\" t1 on t1.\"DocEntry\" = i.\"DocEntry\" ";
            lstrquery += " LEFT JOIN OVTG o ON o.\"Code\" =i.\"U_TaxCode\" ";
            lstrquery += "  where i.\"DocEntry\"='" + DocEntry + "'";
            lstrquery += " GROUP BY \"U_Comp\",o.\"Account\",  ";
            lstrquery += " t1.\"U_DocDate\",t1.\"U_Remarks\" ";
            lstrquery += " UNION all ";
            lstrquery += " SELECT '" + clsModule.objaddon.objcompany.CompanyDB + "' \"U_OffComp\", ";
            lstrquery += " '' \"U_GLCode\",'OthDB' AS types, ";
            lstrquery += " sum(\"U_OffDebit\" -\"U_OffCredit\"),\"U_OffComp\", ";
            lstrquery += " \"U_OffCost1\" ,\"U_OffCost2\" ,\"U_OffCost3\" ,\"U_OffCost4\" ,\"U_OffCost5\",t1.\"U_DocDate\",t1.\"U_Remarks\"   FROM \"@ITC2\" i ";
            lstrquery += " left join \"@OITC\" t1 on t1.\"DocEntry\" = i.\"DocEntry\" ";
            lstrquery += "  where i.\"DocEntry\"='" + DocEntry + "'";
            lstrquery += " GROUP BY \"U_OffComp\",\"U_OffLed\"  ,\"U_OffCost1\" ,\"U_OffCost2\" ,\"U_OffCost3\" ,\"U_OffCost4\" ,\"U_OffCost5\", ";
            lstrquery += " t1.\"U_DocDate\",t1.\"U_Remarks\" ";

            rc = clsModule.objaddon.objglobalmethods.GetmultipleRS(lstrquery);
            if (rc.RecordCount > 0)
            {
                string companyName="" ;
                string Code="";
                string Credit="";
                string Debit="";
                string Dim1="";
                string Dim2="";
                string Dim3="";
                string Dim4="";
                string Dim5="";
                string refdate="";
                string Remark="";
                decimal Totval=0;
                string currcredit = "";
                string currdebit = "";

                Dictionary<string, object> companyData = new Dictionary<string, object>();
                for (int i = 0; i < rc.RecordCount; i++)
                {
                    companyName = rc.Fields.Item("U_DBComp").Value.ToString();
                    Code="";
                    Credit="";
                    Debit="";
                    Dim1="";
                    Dim2="";
                    Dim3="";
                    Dim4="";
                    Dim5="";
                    refdate="";
                    Remark="";

                     currcredit = "";
                     currdebit = "";
                    Totval += stf.CtoD(rc.Fields.Item("Amount").Value);

                    if (stf.CtoD(rc.Fields.Item("Amount").Value.ToString()) < 0)
                    {
                        currcredit = Math.Abs(stf.CtoD(rc.Fields.Item("Amount").Value)).ToString();
                    }
                    else
                    {
                        currdebit = stf.CtoD(rc.Fields.Item("Amount").Value).ToString();
                    }


                    string OffsetLed = rc.Fields.Item("U_GLCode").Value.ToString();

                    if (string.IsNullOrEmpty(rc.Fields.Item("U_GLCode").Value.ToString()))
                    {
                        string lstquery = "SELECT \"U_DBOffset\"  FROM \"@CONFIG2\" c WHERE \"U_DBName1\" ='" + rc.Fields.Item("MainDB").Value.ToString() + "';";
                        OffsetLed = clsModule.objaddon.objglobalmethods.getSingleValue(lstquery);
                    }

                    Code = OffsetLed;
                    Credit = currcredit;
                    Debit = currdebit;
                    Dim1 = rc.Fields.Item("U_Dim1").Value.ToString();
                    Dim2 = rc.Fields.Item("U_Dim2").Value.ToString();
                    Dim3 = rc.Fields.Item("U_Dim3").Value.ToString();
                    Dim4 = rc.Fields.Item("U_Dim4").Value.ToString();
                    Dim5 = rc.Fields.Item("U_Dim5").Value.ToString();
                    refdate = rc.Fields.Item("U_DocDate").Value.ToString();
                    Remark = rc.Fields.Item("U_Remarks").Value.ToString();

                    companyData=new Dictionary<string, object>
                    {
                        {"shortname",Code },
                        {"credit", Credit},
                        {"debit", Debit},
                        {"cost1", Dim1},
                        {"cost2", Dim2},
                        {"cost3", Dim3},
                        {"cost4", Dim4},
                        {"cost5", Dim5},
                        { "refDate",refdate},
                        { "Remark",Remark }

                    };

                    int position = new List<string>(companies.Keys).IndexOf(companyName);
                    if (position == -1)
                    {
                        companies.Add(companyName, new List<Dictionary<string, object>> { companyData });
                        clsModule.objaddon.objapplication.StatusBar.SetText(companyName + "processing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    }
                    else
                    {
                        companies[companyName].Add(companyData);
                    }


                    rc.MoveNext();
                }

                rc = clsModule.objaddon.objglobalmethods.GetmultipleRS("SELECT \"U_CardCode\", \"U_DocDate\", \"U_Remarks\"  FROM \"@OITC\" i where i.\"DocEntry\"='" + DocEntry + "'");
                Debit = "";
                Credit = "";
                if (stf.CtoD(Totval) > 0)
                {
                    Credit = Math.Abs(stf.CtoD(Totval)).ToString();
                }
                else
                {
                    Debit = Math.Abs(stf.CtoD(Totval)).ToString();
                }

                companyData = new Dictionary<string, object>
                    {
                        {"shortname", rc.Fields.Item("U_CardCode").Value.ToString()},
                        {"credit", Credit},
                        {"debit", Debit},
                        {"cost1", ""},
                        {"cost2", ""},
                        {"cost3", ""},
                        {"cost4", ""},
                        {"cost5", ""},
                        { "refDate",rc.Fields.Item("U_DocDate").Value.ToString()},
                        { "Remark",rc.Fields.Item("U_Remarks").Value.ToString() }

                    };
                companies[companyName].Add(companyData);
            }




            foreach (var companyName in companies.Keys)
            {
                if (!string.IsNullOrEmpty(companyName))
                {

                    clsModule.objaddon.objapplication.StatusBar.SetText("Starting Connection " + companyName, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    Transaction.anotherCompany(companyName, out objAnothercompany, out BPCodeCustomer, out BPCodeVendor);
                    PostVoucher(companies[companyName]);
                }
            }
            return true;
        }

        private StaticText StaticText0;
        private StaticText StaticText1;
        private EditText EditText0;
        private EditText EditText1;
        private StaticText StaticText2;
        private EditText EditText2;

        private void Button0_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {

                if (pVal.ActionSuccess)
                {
                    //GeneralService oGeneralService;
                    //GeneralData oGeneralData;
                    //GeneralDataParams oGeneralParams;

                    //oGeneralService = clsModule.objaddon.objcompany.GetCompanyService().GetGeneralService("ATPL_OITC");
                    //oGeneralData = (GeneralData)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);
                    //oGeneralParams = (GeneralDataParams)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);

                    //if (!string.IsNullOrEmpty(DocEntry))
                    //{
                    //    oGeneralParams.SetProperty("DocEntry", DocEntry);
                    //    oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                    //}

                    //oGeneralData.SetProperty("U_DocNum", EditText0.Value);
                    //oGeneralData.SetProperty("U_DocDate", stf.GetDate(EditText1.Value));
                    //oGeneralData.SetProperty("U_Remarks", EditText2.Value);

                    //oGeneralData.Child("ITC1").Add();
                    //int rowcount = 0;
                    //for (int i = 1; i <= Matrix0.VisualRowCount; i++)
                    //{
                    //    clsModule.objaddon.objapplication.StatusBar.SetText("Row Count" + i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    //    if (((SAPbouiCOM.EditText)Matrix0.Columns.Item("DBComp").Cells.Item(i).Specific).String != "")
                    //    {
                    //        if (i > oGeneralData.Child("ITC1").Count)
                    //        {
                    //            oGeneralData.Child("ITC1").Add();
                    //        }
                    //        oGeneralData.Child("ITC1").Item(rowcount).SetProperty("U_DBComp", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("DBComp").Cells.Item(i).Specific).String);
                    //        oGeneralData.Child("ITC1").Item(rowcount).SetProperty("U_GLCode", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("GLCode").Cells.Item(i).Specific).String);
                    //        oGeneralData.Child("ITC1").Item(rowcount).SetProperty("U_GLName", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("GLName").Cells.Item(i).Specific).String);
                    //        oGeneralData.Child("ITC1").Item(rowcount).SetProperty("U_GLAcc", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("GLAcc").Cells.Item(i).Specific).String);
                    //        oGeneralData.Child("ITC1").Item(rowcount).SetProperty("U_Debit", stf.ObjtoStr(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Debit").Cells.Item(i).Specific).String));
                    //        oGeneralData.Child("ITC1").Item(rowcount).SetProperty("U_Credit", stf.ObjtoStr(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Credit").Cells.Item(i).Specific).String));
                    //        oGeneralData.Child("ITC1").Item(rowcount).SetProperty("U_OffComp", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OffComp").Cells.Item(i).Specific).String);
                    //        oGeneralData.Child("ITC1").Item(rowcount).SetProperty("U_OffLed", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("OffLed").Cells.Item(i).Specific).String);
                    //        oGeneralData.Child("ITC1").Item(rowcount).SetProperty("U_Dim1", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Dim1").Cells.Item(i).Specific).String);
                    //        oGeneralData.Child("ITC1").Item(rowcount).SetProperty("U_Dim2", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Dim2").Cells.Item(i).Specific).String);
                    //        oGeneralData.Child("ITC1").Item(rowcount).SetProperty("U_Dim3", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Dim3").Cells.Item(i).Specific).String);
                    //        oGeneralData.Child("ITC1").Item(rowcount).SetProperty("U_Dim4", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Dim4").Cells.Item(i).Specific).String);
                    //        oGeneralData.Child("ITC1").Item(rowcount).SetProperty("U_Dim5", ((SAPbouiCOM.EditText)Matrix0.Columns.Item("Dim5").Cells.Item(i).Specific).String);
                    //        rowcount++;
                    //    }
                    //}
                    //if (!string.IsNullOrEmpty(DocEntry))
                    //{
                    //    oGeneralService.Update(oGeneralData);
                    //}
                    //else
                    //{
                    //    oGeneralParams = oGeneralService.Add(oGeneralData);
                    //    DocEntry = oGeneralParams.GetProperty("DocEntry").ToString();
                    //}


                    //clsModule.objaddon.objapplication.StatusBar.SetText("Processing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    //if (PostotherDB())
                    //{
                    //    clsModule.objaddon.objapplication.StatusBar.SetText("Data Saved Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    //    Cleartext();
                    //}
                }

            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.InnerException.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

            }

        }

        private void Cleartext()
        {
            Matrix0.Clear();
            startInit();
            
        }

        private void Button0_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            string ErrorMsg = "";
            BubbleEvent = DoValidation(ref ErrorMsg);
            if (!string.IsNullOrEmpty(ErrorMsg))
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ErrorMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            
            clsModule.objaddon.objglobalmethods.Removerow(Matrix0, "DBComp");
            clsModule.objaddon.objglobalmethods.Removerow(Matrix1, "MVen" + ((int)colVendor.U_Comp).ToString());
           
        }

        private bool DoValidation(ref string ErrorMsg)
        {
            bool doValidation = true;

            if (string.IsNullOrEmpty(EditText4.Value))
            {
                decimal Credit = Enumerable.Range(1, Matrix0.RowCount).Sum(row =>
                            clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Debit").Cells.Item(row).Specific).Value));

                decimal debit = Enumerable.Range(1, Matrix0.RowCount).Sum(row =>
                      clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Credit").Cells.Item(row).Specific).Value));

                if (Credit - debit != 0)
                {
                    ErrorMsg = "Kindly Check  Amount Total Must be Zero ";
                    doValidation = false;
                }


                if (Credit == 0 || debit == 0)
                {
                    ErrorMsg = "Check Credit and Debit Amount ";
                    doValidation = false;

                }

            }
            return doValidation;
        }

        private void Form_DataAddAfter(ref BusinessObjectInfo pVal)
        {
            if (pVal.ActionSuccess)
            {
                DocEntry = oForm.DataSources.DBDataSources.Item("@OITC").GetValue("DocEntry", 0);
                clsModule.objaddon.objapplication.StatusBar.SetText("Processing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                if (PostotherDB())
                {
                    clsModule.objaddon.objapplication.StatusBar.SetText("Data Saved Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                 
                }
            }

        }

        private EditText EditText3;

        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            Cleartext();

        }

        private Button Button2;

        private void Button2_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
           

        }

        private void Button2_ClickAfter(object sboObject, SBOItemEventArg pVal)
        {

            DocEntry = oForm.DataSources.DBDataSources.Item("@OITC").GetValue("DocEntry", 0);
            clsModule.objaddon.objapplication.StatusBar.SetText("Processing...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

            if (PostotherDB())
            {
                clsModule.objaddon.objapplication.StatusBar.SetText("Data Saved Successfully...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            }
        }

        private Matrix Matrix1;
        private EditText EditText4;
        private StaticText StaticText3;
        private LinkedButton LinkedButton0;

        private void EditText4_ChooseFromListBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {        
                    SAPbouiCOM.ChooseFromList oCFL = oForm.ChooseFromLists.Item("CFL_led");
                    SAPbouiCOM.Conditions oConds;
                    SAPbouiCOM.Condition oCond;
                SAPbouiCOM.Conditions oEmptyConds=null;
                oCFL.SetConditions(oEmptyConds);
                oConds = oCFL.GetConditions();

                    oConds = oCFL.GetConditions();
                    oCond = oConds.Add();
                    oCond.Alias = "CardType";
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCond.CondVal ="S";
                    oCFL.SetConditions(oConds);
            }
            catch (Exception ex)
            {

                
            }

        }

        private StaticText StaticText4;
        private EditText EditText5;

        private void EditText4_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                SAPbouiCOM.ISBOChooseFromListEventArg pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
                if (!(pCFL.SelectedObjects == null))
                {
                    odbHeader.SetValue("U_CardName",0,pCFL.SelectedObjects.Columns.Item("CardName").Cells.Item(0).Value.ToString());
                }


                visiblegrid();
            }
            catch (Exception ex)
            {

                
            }

        }

        private void EditText4_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                visiblegrid();
                odbHeader.SetValue("U_CardName", 0, string.IsNullOrEmpty(EditText4.Value) ? "" : EditText5.Value);
               // oForm.Freeze(true);
               // oForm.Settings.MatrixUID = !string.IsNullOrEmpty(EditText4.Value) ? "Mvendor" : "MJE";
               // oForm.Freeze(false);
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        private void visiblegrid()
        {
         
            Matrix1.Item.Visible = !string.IsNullOrEmpty(EditText4.Value);
            Matrix0.Item.Visible = string.IsNullOrEmpty(EditText4.Value);
            Matrix0.AutoResizeColumns();
            Matrix1.AutoResizeColumns();
           
          
        }


        private void Matrix1_KeyDownBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (pVal.ColUID == "MVen" + ((int)colVendor.U_GLCode).ToString() && pVal.CharPressed == 9)
            {

                if (!string.IsNullOrEmpty(((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_GLCode).ToString()).Cells.Item(pVal.Row).Specific).Value)) return;
                if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_Comp).ToString()).Cells.Item(pVal.Row).Specific).Value)) return;
                choose choose = new choose();
                choose.Retval += Choose_Retval;
                choose.lstrquery = "SELECT  \"AcctName\" as \"Name\" ,\"AcctCode\" as \"code\", \"AcctCode\" as \"CtrlCode\" FROM " + ((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_Comp).ToString()).Cells.Item(pVal.Row).Specific).Value.ToString() + ".OACT  where \"LocManTran\" ='N' and  \"Postable\" ='Y'  AND \"FrozenFor\"='N'  order by \"Name\";";

                Currentrow = pVal.Row;
                choose.Show();
                Matrix1.AddRow();
                offsetcomp = false;
                BubbleEvent = false;
            }

            if (pVal.ColUID == "MVen" + ((int)colVendor.U_OffLed).ToString() && pVal.CharPressed == 9)
            {
                if (string.IsNullOrEmpty(((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_OffComp).ToString()).Cells.Item(pVal.Row).Specific).Value)) return;
                if (!string.IsNullOrEmpty(((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_OffLed).ToString()).Cells.Item(pVal.Row).Specific).Value)) return;
                choose choose = new choose();
                choose.Retval += Choose_Retval;

                choose.lstrquery = "SELECT  \"AcctName\" as \"Name\" ,\"AcctCode\" as \"code\", \"AcctCode\" as \"CtrlCode\" FROM " + ((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_OffComp).ToString()).Cells.Item(pVal.Row).Specific).Value.ToString() + ".OACT  where \"LocManTran\" ='N' and  \"Postable\" ='Y'  AND \"FrozenFor\"='N'  order by \"Name\";";
                Currentrow = pVal.Row;
                choose.Show();
                offsetcomp = true;
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
            if (pVal.ColUID == "MVen" + ((int)colVendor.U_Credit).ToString() && allow)
            {
                if (clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_Debit).ToString()).Cells.Item(pVal.Row).Specific).Value) != 0)
                {
                    BubbleEvent = false;
                }
            }

            if (pVal.ColUID == "MVen" + ((int)colVendor.U_Debit).ToString() && allow)
            {
                if (clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_Credit).ToString()).Cells.Item(pVal.Row).Specific).Value) != 0)
                {
                    BubbleEvent = false;
                }
            }

            if (pVal.ColUID == "MVen" + ((int)colVendor.U_OffCredit).ToString() && allow)
            {
                if (clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_OffDebit).ToString()).Cells.Item(pVal.Row).Specific).Value) != 0)
                {
                    BubbleEvent = false;
                }
            }

            if (pVal.ColUID == "MVen" + ((int)colVendor.U_OffDebit).ToString() && allow)
            {
                if (clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_OffCredit).ToString()).Cells.Item(pVal.Row).Specific).Value) != 0)
                {
                    BubbleEvent = false;
                }
            }

        }

        private void Matrix1_GotFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
            try
            {
                clsModule.objaddon.objapplication.Menus.Item("773").Enabled = true;
                clsModule.objaddon.objapplication.Menus.Item("773").Enabled = oForm.Mode == BoFormMode.fm_ADD_MODE;



                if (pVal.ColUID == "MVen" + ((int)colVendor.U_Credit).ToString())
                {
                    if (clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_Debit).ToString()).Cells.Item(pVal.Row).Specific).Value) != 0)
                    {
                        clsModule.objaddon.objapplication.Menus.Item("773").Enabled = false;
                    }
                }

                if (pVal.ColUID == "MVen" + ((int)colVendor.U_Debit).ToString())
                {
                    if (clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_Credit).ToString()).Cells.Item(pVal.Row).Specific).Value) != 0)
                    {
                        clsModule.objaddon.objapplication.Menus.Item("773").Enabled = false;
                    }
                }

            }
            catch (Exception ex)
            {


            }


        }

        private void Form_DataLoadAfter(ref BusinessObjectInfo pVal)
        {
            //throw new System.NotImplementedException();
           if (pVal.ActionSuccess)
            {
                visiblegrid();
            }

        }

        private void Matrix1_LostFocusAfter(object sboObject, SBOItemEventArg pVal)
        {
          
        }

        private void Matrix1_ChooseFromListAfter(object sboObject, SBOItemEventArg pVal)
        {
           
            if (pVal.ColUID== "MVen" + ((int)colVendor.U_TaxCode).ToString())
            {
                SAPbouiCOM.ISBOChooseFromListEventArg pCFL = (SAPbouiCOM.ISBOChooseFromListEventArg)pVal;
           
                if (!(pCFL.SelectedObjects == null))
                {
                    Matrix1.GetLineData(pVal.Row);
                    ODbvender.SetValue(colVendor.U_TaxCode.ToString(), 0, pCFL.SelectedObjects.Columns.Item("Code").Cells.Item(0).Value.ToString());
                    ODbvender.SetValue(colVendor.U_TaxName.ToString(), 0, pCFL.SelectedObjects.Columns.Item("Name").Cells.Item(0).Value.ToString());
                    ODbvender.SetValue(colVendor.U_TaxRate.ToString(), 0, pCFL.SelectedObjects.Columns.Item("Rate").Cells.Item(0).Value.ToString());
                    Matrix1.SetLineData(pVal.Row);

                    
                    
                }
                ODbvender.Clear();
            }

        }

        private void Matrix1_ValidateAfter(object sboObject, SBOItemEventArg pVal)
        {
            decimal taxval = 0;
            if (pVal.ColUID == "MVen" + ((int)colVendor.U_TaxCode).ToString() ||
                pVal.ColUID == "MVen" + ((int)colVendor.U_Credit).ToString() ||
                pVal.ColUID == "MVen" + ((int)colVendor.U_Debit).ToString()
                )

            {

                taxval += TaxCalculate(stf.CtoD(((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_Credit).ToString()).Cells.Item(pVal.Row).Specific).Value), stf.CtoD(((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_TaxRate).ToString()).Cells.Item(pVal.Row).Specific).Value));
                taxval += TaxCalculate(stf.CtoD(((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_Debit).ToString()).Cells.Item(pVal.Row).Specific).Value), stf.CtoD(((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_TaxRate).ToString()).Cells.Item(pVal.Row).Specific).Value));
                ((SAPbouiCOM.EditText)Matrix1.Columns.Item("MVen" + ((int)colVendor.U_TaxAmt).ToString()).Cells.Item(pVal.Row).Specific).Value = taxval.ToString();
            }
        }

        private decimal TaxCalculate(decimal taxamount,decimal taxRate)
        {
            decimal Tax = 0;

            Tax = taxamount * taxRate / 100;


            return Tax;
        }
    }

}


