using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Xml;
using General.Common;
using SAPbouiCOM;
using SAPbouiCOM.Framework;

namespace ALRedda.Business_Objects
{
    public delegate void Returnvalue(SAPbouiCOM.DataTable sender,int row);
    [FormAttribute("chooseList", "Business_Objects/Choose.b1f")]
    class choose : UserFormBase
    {
        public string lstrquery = "";
        public bool close = false;
        private SAPbouiCOM.Form Form;
        private SAPbouiCOM.DataTable Returnval ;        
        private string curcol;
        private XmlDocument xmlDoc = new XmlDocument();
        public event Returnvalue Retval;

        public choose()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Grid0 = ((SAPbouiCOM.Grid)(this.GetItem("Item_0").Specific));
            this.Grid0.DoubleClickAfter += new SAPbouiCOM._IGridEvents_DoubleClickAfterEventHandler(this.Grid0_DoubleClickAfter);
            this.Grid0.KeyDownAfter += new SAPbouiCOM._IGridEvents_KeyDownAfterEventHandler(this.Grid0_KeyDownAfter);
            this.Grid0.ClickAfter += new SAPbouiCOM._IGridEvents_ClickAfterEventHandler(this.Grid0_ClickAfter);
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_2").Specific));
            this.EditText0.KeyDownAfter += new SAPbouiCOM._IEditTextEvents_KeyDownAfterEventHandler(this.EditText0_KeyDownAfter);
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("choose").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.LoadAfter += new SAPbouiCOM.Framework.FormBase.LoadAfterHandler(this.Form_LoadAfter);
            this.VisibleAfter += new SAPbouiCOM.Framework.FormBase.VisibleAfterHandler(this.Form_VisibleAfter);
            this.KeyDownAfter += new KeyDownAfterHandler(this.Form_KeyDownAfter);

        }

        private SAPbouiCOM.Grid Grid0;

        private void OnCustomInitialize()
        {
            Form = clsModule.objaddon.objapplication.Forms.GetForm("chooseList", 0);
        }

        private void Form_LoadAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            

        }

        private void Grid0_ClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {

            curcol = pVal.ColUID;
        }

        private void Form_VisibleAfter(SAPbouiCOM.SBOItemEventArg pVal)
        {
            if (string.IsNullOrWhiteSpace(lstrquery)) { return; }

            Grid0.DataTable.ExecuteQuery(lstrquery);            
            Grid0.CommonSetting.EnableArrowKey = true;
            Grid0.Rows.SelectedRows.Add(0);

            Grid0.AutoResizeColumns();
            editable();
            Colsetting();

            

        }
        private void editable()
        {
            for (int i = 0; i < Grid0.Columns.Count; i++)
            {
                SAPbouiCOM.GridColumn column = Grid0.Columns.Item(i);
                column.Editable = false;

            }
        }
        private void Colsetting()
        {
            for (int i = 0; i < this.Grid0.Columns.Count; i++)
            {
                this.Grid0.Columns.Item(i).TitleObject.Sortable = true;

            }
            this.Grid0.Columns.Item(0).Editable = false;
        }

        private void Grid0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
           
        }

        private void Grid0_DoubleClickAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
          
            if (pVal.Row == -1) return;
            SelectClose(pVal.Row);


        }

        private void SelectClose(int Row)
        {
            Returnval = Grid0.DataTable ;
            Retval.Invoke(Returnval,Row);
            Form.Close();
        }

        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText0;

        private void EditText0_KeyDownAfter(object sboObject, SAPbouiCOM.SBOItemEventArg pVal)
        {
            int rowIndex;
            if (string.IsNullOrEmpty(curcol))
            {
                curcol = clsModule.objaddon.objglobalmethods.GetColumnValue(Grid0, 0);
            }

            SAPbouiCOM.Grid grid =Grid0;
            switch(pVal.CharPressed)
            {
                case 13:

                    SelectClose(grid.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder));
                    break;
                case 38:
                    if (grid.Rows.SelectedRows.Count > 0)
                    {
                        rowIndex = grid.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder);
                    }
                    else
                    {
                        rowIndex = 0;
                    }

                    if (rowIndex>0)
                    {
                        grid.Rows.SelectedRows.Clear();
                        grid.Rows.SelectedRows.Add(rowIndex-1);
                    }
                    break;

                case 40:
                    if (grid.Rows.SelectedRows.Count > 0)
                    {
                        rowIndex = grid.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder);
                    }
                    else
                    {
                        rowIndex = 0;
                    }

                    if (rowIndex >= 0)
                    {
                        grid.Rows.SelectedRows.Clear();
                        grid.Rows.SelectedRows.Add(rowIndex+1);
                    }
                    break;

                default :
                    rowIndex = Enumerable.Range(0, grid.Rows.Count).FirstOrDefault(index => grid.DataTable.GetValue(curcol, grid.GetDataTableRowIndex(index)).ToString().ToUpper().StartsWith(EditText0.Value.ToString().ToUpper()));
                    if (rowIndex != -1)
                    {
                        grid.Rows.SelectedRows.Clear();
                        grid.Rows.SelectedRows.Add(rowIndex);
                    }
                    break;
            }
            

        }

        private Button Button0;

        private void Button0_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;                   
       SelectClose(Grid0.Rows.SelectedRows.Item(0, BoOrderType.ot_RowOrder));
        }

        private void Form_KeyDownAfter(SBOItemEventArg pVal)
        {
            
        }
    }
}
