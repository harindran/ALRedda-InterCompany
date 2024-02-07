using General.Common;
using SAPbobsCOM;
using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ALRedda.Business_Objects
{
    public static class Transaction
    {


        public static bool anotherCompany(string DBName, out SAPbobsCOM.Company objAnothercompany, out string BPCodeVendor, out string BPCodeCustomer)
        {
            string strSQL = "";
            objAnothercompany = new SAPbobsCOM.Company();
            BPCodeVendor = "";
            BPCodeCustomer = "";

            strSQL = "Select \"U_DBName\",\"U_DBUser\",\"U_DBPass\",\"U_sysUser\",\"U_sysPass\",\"U_BPCode\",\"U_BPCode1\" ";
            strSQL += @" from ""@CONFIG"" T0 join ""@CONFIG1"" T1 on T0.""Code""=T1.""Code"" where T0.""Code""='01' and T1.""U_DBName"" ='" + DBName + "'";

            SAPbobsCOM.Recordset rc = clsModule.objaddon.objglobalmethods.GetmultipleRS(strSQL);
            if (rc.RecordCount > 0)
            {
                objAnothercompany.Server = clsModule.objaddon.objcompany.Server;
                objAnothercompany.LicenseServer = clsModule.objaddon.objcompany.LicenseServer;
                objAnothercompany.SLDServer = clsModule.objaddon.objcompany.SLDServer;
                objAnothercompany.DbServerType = clsModule.objaddon.objcompany.DbServerType;
                objAnothercompany.CompanyDB = DBName;
                objAnothercompany.DbUserName = rc.Fields.Item("U_DBUser").Value.ToString(); //OECDBBR
                objAnothercompany.DbPassword = rc.Fields.Item("U_DBPass").Value.ToString(); //"India@1947";
                objAnothercompany.UserName = rc.Fields.Item("U_sysUser").Value.ToString(); //"tmicloud\\tech.user02";
                objAnothercompany.Password = rc.Fields.Item("U_sysPass").Value.ToString(); //"92W&45KdGsH*";
                objAnothercompany.UseTrusted = false;

                BPCodeVendor = rc.Fields.Item("U_BPCode").Value.ToString();
                BPCodeCustomer = rc.Fields.Item("U_BPCode1").Value.ToString();
                int result = objAnothercompany.Connect();
                clsModule.objaddon.objapplication.StatusBar.SetText(DBName + "processing", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                string error;
                if (result != 0)
                {
                    objAnothercompany.GetLastError(out result, out error);
                    clsModule.objaddon.objapplication.StatusBar.SetText(error, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                    return false;
                }

            }
            return true;
        }


        public static bool cancelJournalandInterreconDocument(SAPbobsCOM.Company _oCompany, int journalDocEntry, int InterRecon)
        {
            SAPbobsCOM.Documents oCancelI;
            SAPbobsCOM.JournalEntries oJV;
            SAPbobsCOM.IInternalReconciliationsService Intrec;
            oJV = (SAPbobsCOM.JournalEntries)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            oCancelI = (SAPbobsCOM.Documents)_oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
            if (!string.IsNullOrEmpty(journalDocEntry.ToString()))
            {
                if (oJV.GetByKey(journalDocEntry))
                {
                    int res = oJV.Cancel();

                    string strerr = "";
                    if (res != 0)
                    {
                        _oCompany.GetLastError(out res, out strerr);
                        clsModule.objaddon.objglobalmethods.WriteErrorLog(strerr);

                    }
                }
            }

            //Intrec = (SAPbobsCOM.InternalReconciliationsService)_oCompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService);
            //InternalReconciliationParams reconParams = (SAPbobsCOM.InternalReconciliationParams)Intrec.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams);
            //if (!string.IsNullOrEmpty(InterRecon.ToString()))
            //{
            //    reconParams.ReconNum = InterRecon;
            //    try
            //    {
            //        Intrec.Cancel(reconParams);
            //    }
            //    catch (Exception ex)
            //    {
            //        clsModule.objaddon.objglobalmethods.WriteErrorLog(ex.Message);
            //        return false;
            //    }
            //}
            return true;

        }

        public static void addLoadMatrixCol(SAPbouiCOM.Matrix Matrix1, string DBName)
        {

            if (Matrix1.RowCount != 0)
            {
                return;

            }

            SAPbobsCOM.Recordset Rcs = clsModule.objaddon.objglobalmethods.GetmultipleRS("SELECT \"DimDesc\"  FROM " + DBName + ".ODIM WHERE \"DimActive\" = 'Y';");
            int i = 0;
            int index = -1;
            do
            {
                string column = "Dim" + (i + 1).ToString();
                index = clsModule.objaddon.objglobalmethods.GetColumnindex(Matrix1, column);
                if (index == -1) break;

                Matrix1.Columns.Remove(index);
                i++;
            } while (true);

            if (Rcs.RecordCount > 0)
            {
                i = 0;
                do
                {

                    string column = "Dim" + (i + 1).ToString();
                    string Description = Rcs.Fields.Item("DimDesc").Value.ToString();
                    Matrix1.Columns.Add(column, BoFormItemTypes.it_EDIT);
                    gridColType(Matrix1, column, Description, Editable: true,
                        width: 50,
                        TableName: "@ODBJE",
                        TableAlias: "U_" + column);
                    i++;
                    Rcs.MoveNext();
                } while (!Rcs.EoF);


            }
        }

        public static void gridColType(SAPbouiCOM.Grid grid, string columnUID, string Description, bool Right = false,
            SAPbouiCOM.BoGridColumnType columnType = SAPbouiCOM.BoGridColumnType.gct_EditText,
            SAPbouiCOM.BoLinkedObject LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_None,
           bool Editable = false, int ForeColor = 0, bool visible = true, int width = 50)
        {
            SAPbouiCOM.EditTextColumn oColumns;
            oColumns = (SAPbouiCOM.EditTextColumn)grid.Columns.Item(columnUID);
            oColumns.LinkedObjectType = Convert.ToInt32(LinkedObjectType).ToString();
            oColumns.Type = columnType;
            oColumns.RightJustified = Right;
            oColumns.Editable = Editable;
            oColumns.ForeColor = ForeColor;
            oColumns.Visible = visible;
            oColumns.Width = width;
            oColumns.Description = Description;
        }
        public static void gridColType(SAPbouiCOM.Matrix grid, string columnUID, string Description, bool Right = false,
            SAPbouiCOM.BoLinkedObject LinkedObjectType = SAPbouiCOM.BoLinkedObject.lf_None,
           bool Editable = false, int ForeColor = 0, bool visible = true, int width = 50, string TableName = "", string TableAlias = "")
        {
            SAPbouiCOM.Column oColumns;
            oColumns = (SAPbouiCOM.Column)grid.Columns.Item(columnUID);
            if (LinkedObjectType != SAPbouiCOM.BoLinkedObject.lf_None)
            {
                LinkedButton oLinkLns = ((SAPbouiCOM.LinkedButton)(oColumns.ExtendedObject));
                oLinkLns.LinkedObject = LinkedObjectType;
            }
            oColumns.RightJustified = Right;
            oColumns.Editable = Editable;
            oColumns.ForeColor = ForeColor;
            oColumns.Visible = visible;
            oColumns.Description = Description;
            oColumns.Width = width;
            oColumns.TitleObject.Caption = Description;
            oColumns.DataBind.SetBound(!string.IsNullOrEmpty(TableName), TableName, TableAlias);

        }

        public static bool reconcile(SAPbobsCOM.Company objAnothercompany, SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix Matrix1, string BPCode, int transid, SAPbobsCOM.BoObjectTypes boObjectTypes)
        {
            try
            {
                SAPbobsCOM.IInternalReconciliationsService obj;
                obj = (SAPbobsCOM.IInternalReconciliationsService)objAnothercompany.GetCompanyService().GetBusinessService(ServiceTypes.InternalReconciliationsService);
                InternalReconciliationOpenTrans openTrans = (InternalReconciliationOpenTrans)obj.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationOpenTrans);
                IInternalReconciliationParams reconParams = (IInternalReconciliationParams)obj.GetDataInterface(InternalReconciliationsServiceDataInterfaces.irsInternalReconciliationParams);
                openTrans.CardOrAccount = CardOrAccountEnum.coaCard;

                SAPbobsCOM.Recordset objRs = (SAPbobsCOM.Recordset)objAnothercompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset objopo = (SAPbobsCOM.Recordset)objAnothercompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                double RecAmount;
                int Row = 0;
                string lstr = "select CASE WHEN T1.\"BalDueCred\"<>0  THEN  T1.\"BalDueCred\" ELSE T1.\"BalDueDeb\" END AS \"Balance\",T1.\"Line_ID\", " +
                                    " CASE WHEN T1.\"BalFcCred\"<>0  THEN  T1.\"BalFcCred\" ELSE T1.\"BalFcDeb\" END AS \"FCBalance\" " +
                             "  from OJDT T0 join JDT1 T1 ON T0.\"TransId\"=T1.\"TransId\" where  T1.\"TransId\"='" + transid + "' and T1.\"ShortName\"='" + BPCode + "'";
                objRs.DoQuery(lstr);
                clsModule.objaddon.objglobalmethods.WriteErrorLog(lstr);
                if (objRs.RecordCount > 0)
                {
                    openTrans.ReconDate = DateTime.Now;
                    for (int Rec = 0; Rec <= objRs.RecordCount - 1; Rec++)
                    {
                        if (Convert.ToDecimal(objRs.Fields.Item("Balance").Value.ToString()) != 0)
                        {
                            if (System.Convert.ToDouble(objRs.Fields.Item("FCBalance").Value.ToString()) != 0)
                            {
                                RecAmount = System.Convert.ToDouble(objRs.Fields.Item("FCBalance").Value.ToString());
                            }
                            else
                            {
                                RecAmount = System.Convert.ToDouble(objRs.Fields.Item("Balance").Value.ToString());
                            }

                            openTrans.InternalReconciliationOpenTransRows.Add();
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES;
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = transid;
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = Convert.ToInt32(objRs.Fields.Item("Line_ID").Value.ToString());
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = RecAmount;
                            Row += 1;
                        }
                        objRs.MoveNext();
                    }
                }

                for (int i = 0; i < Matrix1.RowCount; i++)
                {
                    bool ss2 = ((SAPbouiCOM.CheckBox)Matrix1.Columns.Item("Select").Cells.Item(i + 1).Specific).Checked;
                    if (ss2)
                    {
                        lstr = " select CASE WHEN T1.\"BalDueCred\"<>0  THEN  T1.\"BalDueCred\" ELSE T1.\"BalDueDeb\" END AS \"Balance\",T1.\"Line_ID\",T1.\"TransId\", " +
                            "   ( CASE WHEN  \"FcTotal\" =0 THEN 1 ELSE \"SysTotal\"/\"FcTotal\" END ) AS \"DocRate\" from OJDT T0 join JDT1 T1 ON T0.\"TransId\"=T1.\"TransId\" " +
                            "  where  T1.\"SourceID\"='" + ((SAPbouiCOM.EditText)Matrix1.Columns.Item("DocEntry").Cells.Item(i + 1).Specific).String + "'" +
                            "  and T1.\"ShortName\"='" + BPCode + "'";

                        objRs.DoQuery(lstr);
                        clsModule.objaddon.objglobalmethods.WriteErrorLog(lstr);
                        if (objRs.RecordCount > 0)
                        {
                            RecAmount = Convert.ToDouble(((SAPbouiCOM.EditText)Matrix1.Columns.Item("TotPay").Cells.Item(i + 1).Specific).Value.ToString());
                            double exchange = Convert.ToDouble(objRs.Fields.Item("DocRate").Value.ToString());
                            exchange = 1;
                            openTrans.InternalReconciliationOpenTransRows.Add();
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).Selected = BoYesNoEnum.tYES;
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransId = Convert.ToInt32(objRs.Fields.Item("TransId").Value.ToString());
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).TransRowId = Convert.ToInt32(objRs.Fields.Item("Line_ID").Value.ToString());
                            openTrans.InternalReconciliationOpenTransRows.Item(Row).ReconcileAmount = Math.Round(RecAmount / exchange, Convert.ToInt16(clsModule.objaddon.objglobalmethods.getSingleValue("SELECT Top 1  \"SumDec\"  FROM oADM")));
                            Row += 1;
                        }

                    }
                }

                for (int ik = 0; ik < openTrans.InternalReconciliationOpenTransRows.Count; ik++)
                {
                    clsModule.objaddon.objglobalmethods.WriteErrorLog(openTrans.InternalReconciliationOpenTransRows.Item(ik).TransId.ToString());
                    clsModule.objaddon.objglobalmethods.WriteErrorLog(openTrans.InternalReconciliationOpenTransRows.Item(ik).TransRowId.ToString());
                    clsModule.objaddon.objglobalmethods.WriteErrorLog(openTrans.InternalReconciliationOpenTransRows.Item(ik).ReconcileAmount.ToString());
                }

                try
                {
                    reconParams = obj.Add(openTrans);
                }
                catch (Exception ex)
                {
                    clsModule.objaddon.objglobalmethods.WriteErrorLog(ex.ToString());
                }

                int recnum = reconParams.ReconNum;
                string table = "";
               
                switch (oForm.TypeEx)
                {
                    case "170":
                        table = "ORCT";                        
                        break;
                    case "426":
                        table = "OVPM";                        
                        break;
                }
                savePayUDF(Convert.ToInt32(oForm.DataSources.DBDataSources.Item(table).GetValue("DocEntry", 0)), "U_ReconEnt", recnum.ToString(), boObjectTypes);
                return true;
            }

            catch (Exception ex)
            {
                return false;
            }
        }

        public static bool loadSaveData(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix Matrix1)
        {
            string lstrquery = "";
            string table = "";
            try
            {
                if (oForm == null) return false;
                switch (oForm.TypeEx)
                {
                    case "170":
                        table = "ORCT";
                        break;
                    case "426":
                        table = "OVPM";
                        break;
                }
                string DocEntry = oForm.DataSources.DBDataSources.Item(table).GetValue("DocEntry", 0);
                lstrquery = " SELECT  \"U_Selected\" ,\"U_DocEntry\" ,\"U_DocumentNo\" ,\"U_DocType\" , \"U_BPLId\", " +
                            " \"U_DocDate\" ,\"U_OverDueDay\",\"U_SysDocval\",\"U_FCDocval\" ,\"U_TotAmount\" ,\"U_TotPayment\", " +
                            " \"U_Dim1\" ,\"U_Dim2\",\"U_Dim3\",\"U_Dim4\" ,\"U_Dim5\"  " +
                            "  FROM \"@ODBJE\" WHERE \"U_OGDocEntry\" ='" + DocEntry + "' and \"U_objType\"='" + table + "'";

                SAPbobsCOM.Recordset rc = clsModule.objaddon.objglobalmethods.GetmultipleRS(lstrquery);
                if (rc.RecordCount > 0)
                {
                    Matrix1.Clear();
                    for (int i = 0; i < rc.RecordCount; i++)
                    {
                        Matrix1.AddRow();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("DocEntry").Cells.Item(i + 1).Specific).String = rc.Fields.Item("U_DocEntry").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("DocNo").Cells.Item(i + 1).Specific).String = rc.Fields.Item("U_DocumentNo").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("DocType").Cells.Item(i + 1).Specific).String = rc.Fields.Item("U_DocType").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("BPLId").Cells.Item(i + 1).Specific).String = rc.Fields.Item("U_BPLId").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("DocDate").Cells.Item(i + 1).Specific).String = clsModule.objaddon.objglobalmethods.Getdateformat(rc.Fields.Item("U_DocDate").Value.ToString());
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("OverDD").Cells.Item(i + 1).Specific).String = rc.Fields.Item("U_OverDueDay").Value.ToString();

                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("SysDocval").Cells.Item(i + 1).Specific).String = rc.Fields.Item("U_SysDocval").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("FCDocAmt").Cells.Item(i + 1).Specific).String = rc.Fields.Item("U_FCDocval").Value.ToString();

                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("TotAmt").Cells.Item(i + 1).Specific).String = rc.Fields.Item("U_TotAmount").Value.ToString();
                        ((SAPbouiCOM.EditText)Matrix1.Columns.Item("TotPay").Cells.Item(i + 1).Specific).String = rc.Fields.Item("U_TotPayment").Value.ToString();
                        int col = 0;
                        do
                        {
                            string column = "Dim" + (col + 1).ToString();
                            if (clsModule.objaddon.objglobalmethods.GetColumnindex(Matrix1, column) == -1) break;
                            Matrix1.Columns.Item(column).Editable = false;
                            ((SAPbouiCOM.EditText)Matrix1.Columns.Item(column).Cells.Item(i + 1).Specific).String = rc.Fields.Item("U_" + column).Value.ToString();
                            col++;
                        } while (true);
                        rc.MoveNext();
                        Matrix1.Columns.Item("Select").Visible = false;
                        Matrix1.Columns.Item("TotPay").Editable = false;
                    }
                }
                return true;
            }

            catch (Exception ex)
            {

                return false;
            }
        }

        public static  bool saveTransaction(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix Matrix1)
        {
            GeneralService oGeneralService;
            GeneralData oGeneralData;
            GeneralDataParams oGeneralParams;
            string table = "";
            try
            {
                if (oForm == null) return false;
                switch (oForm.TypeEx)
                {
                    case "170":
                        table = "ORCT";
                        break;
                    case "426":
                        table = "OVPM";
                        break;
                }
            
                oGeneralService = clsModule.objaddon.objcompany.GetCompanyService().GetGeneralService("ATPL_ODBJE");
                oGeneralData = (GeneralData)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);
                for (int i = 0; i < Matrix1.RowCount; i++)
                {

                    bool ss2 = ((SAPbouiCOM.CheckBox)Matrix1.Columns.Item("Select").Cells.Item(i + 1).Specific).Checked;
                    if (ss2)
                    {
                        oGeneralData.SetProperty("U_Selected", "Y");
                        oGeneralData.SetProperty("U_OGDocEntry", oForm.DataSources.DBDataSources.Item(table).GetValue("DocEntry", 0));
                        oGeneralData.SetProperty("U_DocEntry", ((SAPbouiCOM.EditText)Matrix1.Columns.Item("DocEntry").Cells.Item(i + 1).Specific).Value.ToString());
                        oGeneralData.SetProperty("U_DocumentNo", ((SAPbouiCOM.EditText)Matrix1.Columns.Item("DocNo").Cells.Item(i + 1).Specific).Value.ToString());
                        oGeneralData.SetProperty("U_DocType", ((SAPbouiCOM.EditText)Matrix1.Columns.Item("DocType").Cells.Item(i + 1).Specific).Value.ToString());
                        oGeneralData.SetProperty("U_BPLId", ((SAPbouiCOM.EditText)Matrix1.Columns.Item("BPLId").Cells.Item(i + 1).Specific).Value.ToString());
                        oGeneralData.SetProperty("U_DocDate", ((SAPbouiCOM.EditText)Matrix1.Columns.Item("DocDate").Cells.Item(i + 1).Specific).Value.ToString());
                        oGeneralData.SetProperty("U_OverDueDay", ((SAPbouiCOM.EditText)Matrix1.Columns.Item("OverDD").Cells.Item(i + 1).Specific).Value.ToString());
                        oGeneralData.SetProperty("U_SysDocval", ((SAPbouiCOM.EditText)Matrix1.Columns.Item("SysDocval").Cells.Item(i + 1).Specific).Value.ToString());
                        oGeneralData.SetProperty("U_FCDocval", ((SAPbouiCOM.EditText)Matrix1.Columns.Item("FCDocAmt").Cells.Item(i + 1).Specific).Value.ToString());
                        oGeneralData.SetProperty("U_TotAmount", ((SAPbouiCOM.EditText)Matrix1.Columns.Item("TotAmt").Cells.Item(i + 1).Specific).Value.ToString());
                        oGeneralData.SetProperty("U_TotPayment", ((SAPbouiCOM.EditText)Matrix1.Columns.Item("TotPay").Cells.Item(i + 1).Specific).Value.ToString());
                        oGeneralData.SetProperty("U_objType", table);
                        int col = 0;
                        do
                        {
                            string column = "Dim" + (col + 1).ToString();
                            if (clsModule.objaddon.objglobalmethods.GetColumnindex(Matrix1, column) == -1) break;
                            string ss = ((SAPbouiCOM.EditText)Matrix1.Columns.Item(column).Cells.Item(i + 1).Specific).Value.ToString();
                            oGeneralData.SetProperty("U_" + column, ((SAPbouiCOM.EditText)Matrix1.Columns.Item(column).Cells.Item(i + 1).Specific).Value.ToString());
                            col++;
                        } while (true);

                        oGeneralParams = oGeneralService.Add(oGeneralData);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objglobalmethods.WriteErrorLog(ex.ToString());
                return false;
            }

        }

        public static bool savePayUDF(int DocEntry,string UDFColumn, string UDFValue, SAPbobsCOM.BoObjectTypes boObjectTypes )
        {

            SAPbobsCOM.Payments obj = null;
            obj = (SAPbobsCOM.Payments)clsModule.objaddon.objcompany.GetBusinessObject(boObjectTypes);
            if (obj.GetByKey(DocEntry))
            {
                obj.UserFields.Fields.Item(UDFColumn).Value = UDFValue;
                obj.Update();
            }
            return true;
        }

        public static bool saveDOCUDF(int DocEntry, string UDFColumn, string UDFValue, SAPbobsCOM.BoObjectTypes boObjectTypes)
        {

            SAPbobsCOM.Documents obj = null;
            obj = (SAPbobsCOM.Documents)clsModule.objaddon.objcompany.GetBusinessObject(boObjectTypes);
            if (obj.GetByKey(DocEntry))
            {
                obj.UserFields.Fields.Item(UDFColumn).Value = UDFValue;
                obj.Update();
            }
            return true;
        }

        
    }
}
