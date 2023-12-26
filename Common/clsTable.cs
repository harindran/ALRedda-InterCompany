using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbouiCOM.Framework;
namespace General.Common
{
    class clsTable
    {        
        public void FieldCreation()
        {
            AddFields("OVPM", "DBName", "DB Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("OVPM", "CusCode", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("OVPM", "DBbranch", "Branch Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("OVPM", "JouEnt", "Journal Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("OVPM", "ReconEnt", "Reconcile Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);


            AddFields("JDT1", "ODBNames", "ODBName", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("JDT1", "ODBAccNo", "ODB G/L Acct/BP Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);            
            AddFields("JDT1", "ODBAccName", "ODB G/L Acct/BP Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("JDT1", "ODBBranches", "ODB Branch ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("JDT1", "ODBBranchName", "ODB Branch Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("JDT1", "ODBControlAcc", "ODB Control Acc", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("JDT1", "ODBDebit", "ODB Debit", SAPbobsCOM.BoFieldTypes.db_Float,nSubType:SAPbobsCOM.BoFldSubTypes.st_Rate);
            AddFields("JDT1", "ODBCredit", "ODB Credit", SAPbobsCOM.BoFieldTypes.db_Float,nSubType:SAPbobsCOM.BoFldSubTypes.st_Rate);

            AddFields("JDT1", "Dim1", "Dimension 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("JDT1", "Dim2", "Dimension 2", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("JDT1", "Dim3", "Dimension 3", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("JDT1", "Dim4", "Dimension 4", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("JDT1", "Dim5", "Dimension 5", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);

            AddFields("OJDT", "JouEnt", "Journal Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("OJDT", "ReconEnt", "Reconcile Entry", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);




            AddTables("ODBJE", "Other Database", SAPbobsCOM.BoUTBTableType.bott_Document);

            AddFields("@ODBJE", "Selected", "Selected", SAPbobsCOM.BoFieldTypes.db_Alpha, 1);
            AddFields("@ODBJE", "OGDocEntry", "OutGoing DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha,100);
            AddFields("@ODBJE", "DocEntry", "DocEntry", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@ODBJE", "DocumentNo", "Document  Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@ODBJE", "DocType", "Document  Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@ODBJE", "BPLId", "BP ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 10);
            AddFields("@ODBJE", "DocDate", "Document  Date", SAPbobsCOM.BoFieldTypes.db_Date);
            AddFields("@ODBJE", "OverDueDay", "Over Due Days", SAPbobsCOM.BoFieldTypes.db_Numeric);
            AddFields("@ODBJE", "SysDocval", "System Document Value", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@ODBJE", "FCDocval", "FC Document  Value", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);

            AddFields("@ODBJE", "Dim1", "Dimension 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@ODBJE", "Dim2", "Dimension 2", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@ODBJE", "Dim3", "Dimension 3", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@ODBJE", "Dim4", "Dimension 4", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@ODBJE", "Dim5", "Dimension 5", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            

            AddFields("@ODBJE", "objType", "object Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@ODBJE", "TotAmount", "Total Amount", SAPbobsCOM.BoFieldTypes.db_Float,nSubType:SAPbobsCOM.BoFldSubTypes.st_Sum);
            AddFields("@ODBJE", "TotPayment", "Total Payment", SAPbobsCOM.BoFieldTypes.db_Float, nSubType: SAPbobsCOM.BoFldSubTypes.st_Sum);

            AddUDO("ATPL_ODBJE", "Other Database Journal Entry", SAPbobsCOM.BoUDOObjType.boud_Document, "ODBJE", new[] { "" }, new[] { "DocEntry", "DocNum" }, true, true);


            AddTables("OITC", "Inter Company Table", SAPbobsCOM.BoUTBTableType.bott_Document);
            AddTables("ITC1", "Inter Company Table Line", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);


            AddFields("@OITC", "DocNum", "Document Number", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@OITC", "DocDate", "Document Date", SAPbobsCOM.BoFieldTypes.db_Date);            
            AddFields("@OITC", "Remarks", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);            

            AddFields("@ITC1", "DBComp", "DB Company", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@ITC1", "GLCode", "GL Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@ITC1", "GLName", "GL Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@ITC1", "GLAcc", "GL Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@ITC1", "Debit", "Debit", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@ITC1", "Credit", "Credit", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);
            AddFields("@ITC1", "OffComp", "Offset Company", SAPbobsCOM.BoFieldTypes.db_Alpha,50);
            AddFields("@ITC1", "OffLed", "Offset Ledger", SAPbobsCOM.BoFieldTypes.db_Alpha, 50);        
            AddFields("@ITC1", "Dim1", "Dimension 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@ITC1", "Dim2", "Dimension 2", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@ITC1", "Dim3", "Dimension 3", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@ITC1", "Dim4", "Dimension 4", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@ITC1", "Dim5", "Dimension 5", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);

            AddUDO("ATPL_OITC", "Inter Company Table Entry", SAPbobsCOM.BoUDOObjType.boud_Document, "OITC", new[] { "ITC1" }, new[] { "DocEntry", "DocNum" }, true, true);

            #region "Setting Table"
            AddTables("CONFIG", "Configuration Header", SAPbobsCOM.BoUTBTableType.bott_MasterData);
            AddTables("CONFIG1", "Configuration Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);         
            AddTables("CONFIG2", "Configuration2 Lines", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines);         
            AddFields("@CONFIG1", "DBName", "DB Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@CONFIG1", "DBUser", "DB User Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);            
            AddFields("@CONFIG1", "DBPass", "DB Password", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);            
            AddFields("@CONFIG1", "sysUser", "Sys User Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);            
            AddFields("@CONFIG1", "sysPass", "Sys Pass", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);            
            AddFields("@CONFIG1", "BPCode", "BPCode Vendor", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);            
            AddFields("@CONFIG1", "BPCode1", "BPCode Customer", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);

            AddFields("@CONFIG2", "DBName1", "DB Name 1", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@CONFIG2", "DBName2", "DB Name 2", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            AddFields("@CONFIG2", "DBOffset", "DBOffset", SAPbobsCOM.BoFieldTypes.db_Alpha, 100);
            

            AddUDO("ATPL_CONFIG", "Configuration", SAPbobsCOM.BoUDOObjType.boud_MasterData, "CONFIG", new[] { "CONFIG1","CONFIG2" }, new[] { "Code", "Name" }, true, false);
            #endregion "Setting Table"  


        }

        #region Table Creation Common Functions

        private void AddTables(string strTab, string strDesc, SAPbobsCOM.BoUTBTableType nType)
        {
            // var oUserTablesMD = default(SAPbobsCOM.UserTablesMD);
            SAPbobsCOM.UserTablesMD oUserTablesMD = null;
            try
            {
                oUserTablesMD = (SAPbobsCOM.UserTablesMD)clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                // Adding Table
                if (!oUserTablesMD.GetByKey(strTab))
                {
                    oUserTablesMD.TableName = strTab;
                    oUserTablesMD.TableDescription = strDesc;
                    oUserTablesMD.TableType = nType;

                    if (oUserTablesMD.Add() != 0)
                    {
                        throw new Exception(clsModule.objaddon.objcompany.GetLastErrorDescription() + strTab);
                    }
                }
            }
            catch (Exception ex)
            {
                return;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTablesMD);
                oUserTablesMD = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private void AddFields(string strTab, string strCol, string strDesc, SAPbobsCOM.BoFieldTypes nType, int nEditSize = 10, SAPbobsCOM.BoFldSubTypes nSubType = 0, SAPbobsCOM.BoYesNoEnum Mandatory = SAPbobsCOM.BoYesNoEnum.tNO, string defaultvalue = "", bool Yesno = false, string[] Validvalues = null)
        {
            SAPbobsCOM.UserFieldsMD oUserFieldMD1;
            oUserFieldMD1 = (SAPbobsCOM.UserFieldsMD)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
            try
            {
             
                if (!IsColumnExists(strTab, strCol))
                {                   
                    oUserFieldMD1.Description = strDesc;
                    oUserFieldMD1.Name = strCol;
                    oUserFieldMD1.Type = nType;
                    oUserFieldMD1.SubType = nSubType;
                    oUserFieldMD1.TableName = strTab;
                    oUserFieldMD1.EditSize = nEditSize;
                    oUserFieldMD1.Mandatory = Mandatory;
                    oUserFieldMD1.DefaultValue = defaultvalue;

                    if (Yesno == true)
                    {
                        oUserFieldMD1.ValidValues.Value = "Y";
                        oUserFieldMD1.ValidValues.Description = "Yes";
                        oUserFieldMD1.ValidValues.Add();
                        oUserFieldMD1.ValidValues.Value = "N";
                        oUserFieldMD1.ValidValues.Description = "No";
                        oUserFieldMD1.ValidValues.Add();
                    }
                    //if (LinkedSystemObject != 0)
                    //    oUserFieldMD1.LinkedSystemObject = LinkedSystemObject;

                    string[] split_char;
                    if (Validvalues !=null)
            {
                        if (Validvalues.Length > 0)
                        {
                            for (int i = 0, loopTo = Validvalues.Length - 1; i <= loopTo; i++)
                            {
                                if (string.IsNullOrEmpty(Validvalues[i]))
                                    continue;
                                split_char = Validvalues[i].Split(Convert.ToChar(","));
                                if (split_char.Length != 2)
                                    continue;
                                oUserFieldMD1.ValidValues.Value = split_char[0];
                                oUserFieldMD1.ValidValues.Description = split_char[1];
                                oUserFieldMD1.ValidValues.Add();
                            }
                        }
                    }
                    int val;
                    val = oUserFieldMD1.Add();
                    if (val != 0)
                    {
                        clsModule.objaddon.objapplication.SetStatusBarMessage(clsModule. objaddon.objcompany.GetLastErrorDescription() + " " + strTab + " " + strCol, SAPbouiCOM.BoMessageTime.bmt_Short,true);
                    }
                    // System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1)
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldMD1);
                oUserFieldMD1 = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private bool IsColumnExists(string Table, string Column)
        {
            SAPbobsCOM.Recordset oRecordSet=null;
            string strSQL;
            try
            {
               
                strSQL = "SELECT COUNT(*) FROM CUFD WHERE \"TableID\" = '" + Table + "' AND \"AliasID\" = '" + Column + "'";
                             
                oRecordSet = (SAPbobsCOM.Recordset)clsModule.objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRecordSet.DoQuery(strSQL);

                if (Convert.ToInt32( oRecordSet.Fields.Item(0).Value) == 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRecordSet);
                oRecordSet = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }

        private void AddKey(string strTab, string strColumn, string strKey, int i)
        {
            var oUserKeysMD = default(SAPbobsCOM.UserKeysMD);

            try
            {
                // // The meta-data object must be initialized with a
                // // regular UserKeys object
                oUserKeysMD =(SAPbobsCOM.UserKeysMD)clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);

                if (!oUserKeysMD.GetByKey("@" + strTab, i))
                {

                    // // Set the table name and the key name
                    oUserKeysMD.TableName = strTab;
                    oUserKeysMD.KeyName = strKey;

                    // // Set the column's alias
                    oUserKeysMD.Elements.ColumnAlias = strColumn;
                    oUserKeysMD.Elements.Add();
                    oUserKeysMD.Elements.ColumnAlias = "RentFac";

                    // // Determine whether the key is unique or not
                    oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES;

                    // // Add the key
                    if (oUserKeysMD.Add() != 0)
                    {
                        throw new Exception(clsModule. objaddon.objcompany.GetLastErrorDescription());
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD);
                oUserKeysMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void AddUDO(string strUDO, string strUDODesc, SAPbobsCOM.BoUDOObjType nObjectType, string strTable, string[] childTable, string[] sFind, bool canlog = false, bool Manageseries = false)
        {

           SAPbobsCOM.UserObjectsMD oUserObjectMD=null;
            int tablecount = 0;
            try
            {
                oUserObjectMD = (SAPbobsCOM.UserObjectsMD)clsModule. objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
                oUserObjectMD.GetByKey(strUDO);

                if (!oUserObjectMD.GetByKey(strUDO)) //(oUserObjectMD.GetByKey(strUDO) == 0)
                {
                    oUserObjectMD.Code = strUDO;
                    oUserObjectMD.Name = strUDODesc;
                    oUserObjectMD.ObjectType = nObjectType;
                    oUserObjectMD.TableName = strTable;

                    oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES;

                    if (Manageseries)
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                    else
                        oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;

                    if (canlog)
                    {
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES;
                        oUserObjectMD.LogTableName = "A" + strTable.ToString();
                    }
                    else
                    {
                        oUserObjectMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO;
                        oUserObjectMD.LogTableName = "";
                    }

                    oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;
                    oUserObjectMD.ExtensionName = "";

                    oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                    tablecount = 1;
                    if (sFind.Length > 0)
                    {
                        for (int i = 0, loopTo = sFind.Length - 1; i <= loopTo; i++)
                        {
                            if (string.IsNullOrEmpty(sFind[i]))
                                continue;
                            oUserObjectMD.FindColumns.ColumnAlias = sFind[i];
                            oUserObjectMD.FindColumns.Add();
                            oUserObjectMD.FindColumns.SetCurrentLine(tablecount);
                            tablecount = tablecount + 1;
                        }
                    }

                    tablecount = 0;
                    if (childTable != null)
                     {
                        if (childTable.Length > 0)
                        {
                            for (int i = 0, loopTo1 = childTable.Length - 1; i <= loopTo1; i++)
                            {
                                if (string.IsNullOrEmpty(childTable[i]))
                                    continue;
                                oUserObjectMD.ChildTables.SetCurrentLine(tablecount);
                                oUserObjectMD.ChildTables.TableName = childTable[i];
                                oUserObjectMD.ChildTables.Add();
                                tablecount = tablecount + 1;
                            }
                        }
                    }

                    if (oUserObjectMD.Add() != 0)
                    {
                        throw new Exception(clsModule. objaddon.objcompany.GetLastErrorDescription());
                    }
                }

                else
                {
                    tablecount = 0;
                    if (childTable.Length != oUserObjectMD.ChildTables.Count) {
                        if (childTable != null)
                        {
                            if (childTable.Length > 0)
                            {
                                for (int i = 0, loopTo1 = childTable.Length - 1; i <= loopTo1; i++)
                                {
                                    if (string.IsNullOrEmpty(childTable[i]))
                                        continue;
                                    oUserObjectMD.ChildTables.SetCurrentLine(tablecount);
                                    oUserObjectMD.ChildTables.TableName = childTable[i];
                                    oUserObjectMD.ChildTables.Add();
                                    tablecount = tablecount + 1;
                                }
                                if (tablecount > 0)
                                {
                                    oUserObjectMD.Update();
                                }
                            }
                        }
                    }

                }
            }

            catch (Exception ex)
            {
                return;
            }
            finally
            {
                if (oUserObjectMD != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                    oUserObjectMD = null;
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                }
            }

        }


        #endregion


        

    }
}
