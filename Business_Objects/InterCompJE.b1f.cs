using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using General.Common;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using static General.Common.Module;

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

        public InterCompJE()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_0").Specific));
            this.Matrix0.GotFocusAfter += new SAPbouiCOM._IMatrixEvents_GotFocusAfterEventHandler(this.Matrix0_GotFocusAfter);
            this.Matrix0.KeyDownBefore += new SAPbouiCOM._IMatrixEvents_KeyDownBeforeEventHandler(this.Matrix0_KeyDownBefore);
            this.Folder0 = ((SAPbouiCOM.Folder)(this.GetItem("Item_2").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
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
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.DataAddAfter += new DataAddAfterHandler(this.Form_DataAddAfter);

        }

        private SAPbouiCOM.Matrix Matrix0;



        private void OnCustomInitialize()
        {
            startInit();

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
            if (Row >= 0)
            {
                ((SAPbouiCOM.EditText)Matrix0.Columns.Item("GLCode").Cells.Item(Currentrow).Specific).Value = sender.GetValue("code", Row).ToString();
                ((SAPbouiCOM.EditText)Matrix0.Columns.Item("GLName").Cells.Item(Currentrow).Specific).Value = sender.GetValue("Name", Row).ToString();
                ((SAPbouiCOM.EditText)Matrix0.Columns.Item("GLAcc").Cells.Item(Currentrow).Specific).Value = sender.GetValue("CtrlCode", Row).ToString();
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

                oJV.ReferenceDate = stf.GetDate(companies[i].ReferenceDate.ToString());
                oJV.Memo = companies[i].Remark.ToString();

                oJV.Lines.ShortName = companies[i].GroupKey.ToString();
                oJV.Lines.Credit = clsModule.objaddon.objglobalmethods.Cton(companies[i].CreditSum);
                oJV.Lines.Debit = clsModule.objaddon.objglobalmethods.Cton(companies[i].DebitSum);
                oJV.Lines.BPLID = 1;
                oJV.Lines.CostingCode = companies[i].cost1.ToString();
                oJV.Lines.CostingCode2 = companies[i].cost2.ToString();
                oJV.Lines.CostingCode3 = companies[i].cost3.ToString();
                oJV.Lines.CostingCode4 = companies[i].cost4.ToString();
                oJV.Lines.CostingCode5 = companies[i].cost5.ToString();

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
            }

            return true;
        }


        private Button Button0;
        private Button Button1;


        private bool PostotherDB()
        {
            string lstrquery = "";

            lstrquery += "  SELECT \"U_DBComp\" ,\"U_GLCode\" ,\"U_GLName\" ,\"U_GLAcc\" ,\"U_Debit\" ,\"U_Credit\" ,\"U_OffComp\" ,\"U_OffLed\", ";
            lstrquery += "  \"U_Dim1\" ,\"U_Dim2\" ,\"U_Dim3\" ,\"U_Dim4\" ,\"U_Dim5\",\"U_DocDate\",\"U_Remarks\"  FROM \"@OITC\" t1 ";
            lstrquery += "  LEFT JOIN \"@ITC1\" t2  ON t1.\"DocEntry\" =t2.\"DocEntry\" where t1.\"DocEntry\"='" + DocEntry + "' ; ";
            SAPbobsCOM.Recordset rc = clsModule.objaddon.objglobalmethods.GetmultipleRS(lstrquery);
            if (rc.RecordCount > 0)
            {
                Dictionary<string, List<Dictionary<string, object>>> companies = new Dictionary<string, List<Dictionary<string, object>>>();

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

                        if (string.IsNullOrEmpty( rc.Fields.Item("U_OffLed").Value.ToString()) && j!=0)
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





                foreach (var companyName in companies.Keys)
                {
                    if (!string.IsNullOrEmpty(companyName))
                    {

                        clsModule.objaddon.objapplication.StatusBar.SetText("Starting Connection", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                        Transaction.anotherCompany(companyName, out objAnothercompany, out BPCodeCustomer, out BPCodeVendor);
                        PostVoucher(companies[companyName]);
                    }
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
            OnCustomInitialize();
            EditText2.Value = "";
        }

        private void Button0_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            string ErrorMsg = "";
            BubbleEvent = DoValidation(ref ErrorMsg);
            if (!string.IsNullOrEmpty(ErrorMsg))
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ErrorMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            clsModule.objaddon.objglobalmethods.RemoveLastrow(Matrix0, "DBComp");
           
        }

        private bool DoValidation(ref string ErrorMsg)
        {
            bool doValidation = true;


            decimal Credit = Enumerable.Range(1, Matrix0.RowCount).Sum(row =>
                        clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Debit").Cells.Item(row).Specific).Value));

            decimal debit = Enumerable.Range(1, Matrix0.RowCount).Sum(row =>
                  clsModule.objaddon.objglobalmethods.CtoD(((SAPbouiCOM.EditText)Matrix0.Columns.Item("Credit").Cells.Item(row).Specific).Value));

            if (Credit - debit != 0)
            {
                ErrorMsg = "Kindly Check  Amount Total Must be Zero ";
                doValidation = false;
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
                    Cleartext();
                }
            }

        }

        private EditText EditText3;
    }

}


