﻿using ALRedda.Business_Objects;
using Logic.Business_Objects;
using SAPbouiCOM.Framework;
using System;
using System.Data;
using System.IO;

namespace General.Common
{
    class clsAddon
    {
        public clsMenuEvent objmenuevent;
        public clsGlobalMethods objglobalmethods;
        public SAPbouiCOM.Application objapplication;
        public SAPbobsCOM.Company objcompany;
        public clsRightClickEvent objrightclickevent;
        

        public string[] HWKEY = { "L1653539483", "X1211807750","K1600107675" };
        #region Constructor
        public clsAddon()
        {

        }
        #endregion

        public void Intialize(string[] args)
        {
            try
            {
                Application oapplication;
                if ((args.Length < 1))
                    oapplication = new Application();
                else
                    oapplication = new Application(args[0]);
                objapplication = Application.SBO_Application;

              

              
                if (isValidLicense())
                {
                    objapplication.StatusBar.SetText("Establishing Company Connection Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                    objcompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();

                    Create_DatabaseFields(); // UDF & UDO Creation Part    
                    Menu(); // Menu Creation Part
                    Create_Objects(); // Object Creation Part

                   // Add_Authorizations();
                   // SetFilters();
                    objapplication.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(objapplication_AppEvent);
                    objapplication.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(objapplication_MenuEvent);
                    objapplication.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(objapplication_ItemEvent);               
                    objapplication.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(FormDataEvent);
                    objapplication.UDOEvent += new SAPbouiCOM._IApplicationEvents_UDOEventEventHandler(FormUDOEvent);
                    objapplication.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(objapplication_RightClickEvent);                    
                                        
                    

                    objapplication.StatusBar.SetText("Addon Connected Successfully..!!!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                    oapplication.Run();
                }
                else
                {
                    objapplication.StatusBar.SetText("Addon Disconnected due to license mismatch..!!!", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return;
                }
            }          
            catch (Exception ex)
            {
                
                WriteErrorLog(ex.Message);
                objapplication.SetStatusBarMessage(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public void WriteErrorLog(string Str)
        {
            try
            {

                string Foldername;
                Foldername = @"Log";
                if (Directory.Exists(Foldername))
                {
                }
                else
                {
                    Directory.CreateDirectory(Foldername);
                }

                FileStream fs;
                string chatlog = Foldername + @"\Log_" + DateTime.Now.ToString("ddMMyy") + ".txt";
                if (File.Exists(chatlog))
                {
                }
                else
                {
                    fs = new FileStream(chatlog, FileMode.Create, FileAccess.Write);
                    fs.Close();
                }
                string sdate;
                sdate = Convert.ToString(DateTime.Now);
                if (File.Exists(chatlog) == true)
                {
                    var objWriter = new StreamWriter(chatlog, true);
                    objWriter.WriteLine(sdate + " : " + Str);
                    objWriter.Close();
                }
                else
                {
                    var objWriter = new StreamWriter(chatlog, false);
                }
            }
            catch (Exception)
            {


            }
        }
        private void SetFilters()
        {
              SAPbouiCOM.EventFilters oFilters;
         SAPbouiCOM.EventFilter oFilter;

        oFilters = new SAPbouiCOM.EventFilters();

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
            oFilter.AddEx("133");
            oFilter.AddEx("179");                      
                                


            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD);
            oFilter.AddEx("133");
            oFilter.AddEx("179");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE);
            oFilter.AddEx("133");
            oFilter.AddEx("179");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_LOAD);
            oFilter.AddEx("133");
            oFilter.AddEx("179");
            

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS);
            oFilter.AddEx("133");
            oFilter.AddEx("179");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);
            oFilter.AddEx("133");
            oFilter.AddEx("179");


            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DRAW);
            oFilter.AddEx("133");
            oFilter.AddEx("179");

            oFilter = oFilters.Add(SAPbouiCOM.BoEventTypes.et_ALL_EVENTS);
            oFilter.AddEx("EINVMUL");
            oFilter.AddEx("EINVCON");
            oFilter.AddEx("EILOGVIEW");

            objapplication.SetFilter(oFilters);

        }
        public bool isValidLicense()
        {
            return true;
            try
            {
                if (clsModule.HANA)
                {
                    try
                    {
                        if (objapplication.Forms.ActiveForm.TypeCount > 0)
                        {
                            for (int i = 0; i <= objapplication.Forms.ActiveForm.TypeCount - 1; i++)
                                objapplication.Forms.ActiveForm.Close();
                        }
                    }
                    catch (Exception ex)
                    {
                    }
                }

                objapplication.Menus.Item("257").Activate();
                SAPbouiCOM.EditText objedit = (SAPbouiCOM.EditText)objapplication.Forms.ActiveForm.Items.Item("79").Specific;

                string CrrHWKEY = objedit.Value.ToString();
                objapplication.Forms.ActiveForm.Close();

                for (int i = 0; i <= HWKEY.Length - 1; i++)
                {
                    if (HWKEY[i] == CrrHWKEY)
                    {
                        return true;
                    }

                }

                System.Windows.Forms.MessageBox.Show("Installing Add-On failed due to License mismatch");
                return false;
            }
            catch (Exception ex)
            {
                objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            }
            return true;
        }

        public void Create_Objects()
        {
            objmenuevent = new clsMenuEvent();
            objrightclickevent = new clsRightClickEvent();
            objglobalmethods = new clsGlobalMethods();                     
        }

        private void Create_DatabaseFields()
        {
            objapplication.StatusBar.SetText("Creating Database Fields.Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            var objtable = new clsTable();
            objtable.FieldCreation();
            objapplication.StatusBar.SetText(" Database Created Successfully...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
        }

        #region Menu Creation Details

        private void Menu()
        {
      
            if (objapplication.Menus.Item("43520").SubMenus.Exists("Inter Company"))
                return;
            CreateMenu("", objapplication.Menus.Item("43520").SubMenus.Count + 1, "Inter Company", SAPbouiCOM.BoMenuType.mt_POPUP, "Inter Company", "43520");
            CreateMenu("", objapplication.Menus.Item("Inter Company").SubMenus.Count + 1, "Inter Company Transaction", SAPbouiCOM.BoMenuType.mt_STRING, "ICT", "Inter Company");
            CreateMenu("", objapplication.Menus.Item("Inter Company").SubMenus.Count + 1, "DB Configuration", SAPbouiCOM.BoMenuType.mt_STRING, "CONFIG", "Inter Company");
            

        }

        public void Add_Authorizations()
        {
            try
            {

                

                clsModule.objaddon.objglobalmethods.AddToPermissionTree("Altrocks Tech", "ATPL_ADD-ON", "", "", 'N');
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("E-invoice", "ATPL_EINV", "", "ATPL_ADD-ON", 'N');
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("E-invoice Configuration", "ATPL_EINVCON", "EINVCON", "ATPL_EINV", 'N');
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("E-invoice Log", "ATPL_EILOGVIEW", "EILOGVIEW", "ATPL_EINV", 'N');
                clsModule.objaddon.objglobalmethods.AddToPermissionTree("E-invoice Multiple", "ATPL_EINVMUL", "EINVMUL", "ATPL_EINV", 'N');
           


            }
            catch (Exception ex)
            {

            }
        }

        private void CreateMenu(string ImagePath, int Position, string DisplayName, SAPbouiCOM.BoMenuType MenuType, string UniqueID, string ParentMenuID)
        {
            try
            {
                SAPbouiCOM.MenuCreationParams oMenuPackage;
                SAPbouiCOM.MenuItem parentmenu;
                parentmenu = objapplication.Menus.Item(ParentMenuID);
                if (parentmenu.SubMenus.Exists(UniqueID.ToString()))
                    return;
                oMenuPackage = (SAPbouiCOM.MenuCreationParams)objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oMenuPackage.Image = ImagePath;
                oMenuPackage.Position = Position;
                oMenuPackage.Type = MenuType;
                oMenuPackage.UniqueID = UniqueID;
                oMenuPackage.String = DisplayName;
                parentmenu.SubMenus.AddEx(oMenuPackage);
            }
            catch (Exception ex)
            {
                objapplication.StatusBar.SetText("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
            }
        }

        #endregion

        public bool FormExist(string FormID)
        {
            bool FormExistRet = false;
            try
            {
                FormExistRet = false;
                foreach (SAPbouiCOM.Form uid in clsModule.objaddon.objapplication.Forms)
                {
                    if (uid.TypeEx == FormID)
                    {
                        FormExistRet = true;
                        break;
                    }
                }
                if (FormExistRet)
                {
                    clsModule.objaddon.objapplication.Forms.Item(FormID).Visible = true;
                    clsModule.objaddon.objapplication.Forms.Item(FormID).Select();
                }
            }
            catch (Exception ex)
            {
                clsModule.objaddon.objapplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

            return FormExistRet;

        }

       

        #region ItemEvent

        private void objapplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            
            BubbleEvent = true;
         
            clsModule.objaddon.objglobalmethods.WriteErrorLog(pVal.EventType.ToString() + pVal.ActionSuccess + pVal.FormTypeEx+pVal.InnerEvent);
           
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:                      
                        break;

                    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:                       
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:                                               
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:                                            
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DRAW:                       
                        break;                        
                }
              
            }
            catch (Exception ex)
            {
                return;
            }
        }

        #endregion

        private void FormUDOEvent(ref SAPbouiCOM.UDOEvent BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }
            #region FormDataEvent

            private void FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:                        
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:                        
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:                     
                        break;
                }
                
            }
            catch (Exception ex)
            {
               //throw ex;
            }
        }

        #endregion

        #region MenuEvent
        private void objapplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                

           
            if (pVal.BeforeAction == false)
                {
                    switch (pVal.MenuUID)
                    {
                        case "1281":
                        case "1282":
                        case "1283":
                        case "1284":
                        case "1285":
                        case "1286":
                        case "1287":
                        case "1300":
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                        case "1304":
                        case "1292":
                        case "1293":
                            objmenuevent.MenuEvent_For_StandardMenu(ref pVal, ref BubbleEvent);
                            break;

                        case "CONFIG":
                            CONFIG Configuration = new CONFIG();
                            Configuration.Show();
                            break;                        
                        case "ICT":
                            InterCompJE InterCompJE = new InterCompJE();
                            InterCompJE.Show();
                            break;
                    }
                }
            else
                {
                    switch (pVal.MenuUID)
                    {
                        case "1284":
                            bool valid =BubbleEvent;
                            SAPbouiCOM.Form currentfrm = clsModule.objaddon.objapplication.Forms.ActiveForm;
                            
                            switch (currentfrm.TypeEx)
                            {
                                case "170":
                                case "426":                                                                                               
                                    
                                    SAPbobsCOM.Company company;
                                    string BPcodevendor;
                                    string BPcodeCustomer;
                                    string Table="";
                                    switch (currentfrm.TypeEx)
                                    {
                                        case "170":
                                            Table = "ORCT";
                                            break;
                                        case "426":
                                            Table = "OVPM";
                                            break;
                                       
                                    }
                                    currentfrm.DataSources.DBDataSources.Item(Table).GetValue("DocEntry", 0);

                                    
                                    valid = Transaction.anotherCompany(currentfrm.DataSources.DBDataSources.Item(Table).GetValue("U_DBName", 0), out company, out BPcodevendor, out BPcodeCustomer);
                                    valid = Transaction.cancelJournalandInterreconDocument(company, Convert.ToInt16(currentfrm.DataSources.DBDataSources.Item(Table).GetValue("U_JouEnt", 0)), Convert.ToInt16(currentfrm.DataSources.DBDataSources.Item(Table).GetValue("U_ReconEnt", 0)));

                                    break;
                               
                            }

                            switch (currentfrm.TypeEx)
                            {
                                case "392":
                                clsModule.objaddon.objglobalmethods.cancel = true;
                                    break;
                            }

                            BubbleEvent = valid;
                            break;
                    }
                    

                }
            }
            catch (Exception ex)
            {

              
            }
        }

        #endregion
       
        #region RightClickEvent

        private void objapplication_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                switch (objapplication.Forms.ActiveForm.TypeEx)
                {
                    case "138":
                        objrightclickevent.RightClickEvent(ref eventInfo, ref BubbleEvent);
                        break;
                    case "UOMMAP":
                        objrightclickevent.RightClickEvent(ref eventInfo, ref BubbleEvent);
                        break;
                }

            }
            catch (Exception ex) { }

        }

        #endregion

        #region AppEvent

        private void objapplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:                
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    try
                    {
                        System.Windows.Forms.Application.Exit();
                        if (objapplication != null)
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objapplication);
                        if (objcompany != null)
                        {
                            if (objcompany.Connected)
                                objcompany.Disconnect();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompany);
                        }
                        GC.Collect();

                    }
                    catch (Exception ex)
                    {
                    }

                    break;
            }
        }

        #endregion


        private void DisConnect_Addon()
        {
            try
            {
               if (clsModule.objaddon.objapplication.Forms.Count > 0)
                {
                    try
                    {
                        for (int frm = clsModule.objaddon.objapplication.Forms.Count - 1; frm >= 0; frm--)
                        {
                            if (clsModule.objaddon.objapplication.Forms.Item(frm).IsSystem == true) continue;
                            clsModule.objaddon.objapplication.Forms.Item(frm).Close();
                        }
                    }
                    catch (Exception ex)
                    {

                    }
                }
                if (objcompany.Connected)
                    objcompany.Disconnect();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objcompany);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objapplication);
               objcompany = null;
                GC.Collect();
                System.Windows.Forms.Application.Exit();
                System.Environment.Exit(0);
            }
            catch (Exception ex)
            {

            }
        }
        private void Remove_Menu(string[] MenuID = null)
        {
            try
               {
                string[] split_char;
                if (MenuID != null)
                {
                    if (MenuID.Length > 0)
                    {
                        for (int i = 0, loopTo = MenuID.Length - 1; i <= loopTo; i++)
                           {

                            if (string.IsNullOrEmpty(MenuID[i]))

                                continue;

                            split_char = MenuID[i].Split(Convert.ToChar(","));

                            if (split_char.Length != 2)
                                continue;

                            if (clsModule.objaddon.objapplication.Menus.Item(split_char[0]).SubMenus.Exists(split_char[1]))
                                clsModule.objaddon.objapplication.Menus.Item(split_char[0]).SubMenus.RemoveEx(split_char[1]);

                        }

                    }

                }

            }

            catch (Exception ex)

            {

            }



        }

    }


}
