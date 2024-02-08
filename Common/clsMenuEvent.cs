using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ALRedda.Business_Objects.SupportFiles;
namespace General.Common
{
    class clsMenuEvent
    {
        SAPbouiCOM.Form objform;
        SAPbouiCOM.Form oUDFForm;

        public void MenuEvent_For_StandardMenu(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (!pVal.BeforeAction)
                {

                    switch (clsModule.objaddon.objapplication.Forms.ActiveForm.TypeEx)
                    {
                        case "ICT":
                            ICT_MenuEvent(ref pVal, ref BubbleEvent);
                            break;

                    }
                }
                else
                {
                    switch (clsModule.objaddon.objapplication.Forms.ActiveForm.TypeEx)
                    {

                    }
                }
            }
            catch (Exception ex)
            {

            }
        }
        private void ICT_MenuEvent(ref SAPbouiCOM.MenuEvent pval, ref bool BubbleEvent)
        {
            try
            {

                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                if (!pval.BeforeAction)
                {
                    switch (pval.MenuUID)
                    {
                        case "1288":
                        case "1289":
                        case "1290":
                        case "1291":
                        case "1304":
                        case "1282":
                            ITCclscs cNCReq = new ITCclscs(objform);
                            cNCReq.ITCStart();
                            break;
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        private void Default_Sample_MenuEvent(SAPbouiCOM.MenuEvent pval, bool BubbleEvent)
        {
            try
            {
                objform = clsModule.objaddon.objapplication.Forms.ActiveForm;
                if (pval.BeforeAction == true)
                {
                }

                else
                {
                    SAPbouiCOM.Form oUDFForm;
                    try
                    {
                        oUDFForm = clsModule.objaddon.objapplication.Forms.Item(objform.UDFFormUID);
                    }
                    catch (Exception ex)
                    {
                        oUDFForm = objform;
                    }

                    switch (pval.MenuUID)
                    {
                        case "1281": // Find
                            {
                               // oUDFForm.Items.Item("U_IRNNo").Enabled = true;                               
                                break;
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
