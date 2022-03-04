using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;

namespace TestMatrix
{
    [FormAttribute("TestMatrix.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("Item_1").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("btnFill").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.Matrix Matrix1;
        private SAPbouiCOM.Button Button0;

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            //throw new System.NotImplementedException
            SAPbobsCOM.Company oCompany = (SAPbobsCOM.Company)Application.SBO_Application.Company.GetDICompany();
            SAPbobsCOM.Recordset orset = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            String Query = "Select CardCode,CardName,E_Mail From  OCRD ";
            orset.DoQuery(Query);
            if (orset.RecordCount> 0) {

                for (int i=0;i<orset.RecordCount;i++) {

                    Matrix1.AddRow();
                    ((SAPbouiCOM.EditText)Matrix1.Columns.Item("colCode").Cells.Item(i + 1).Specific).Value = orset.Fields.Item("CardCode").Value.ToString();
                    ((SAPbouiCOM.EditText)Matrix1.Columns.Item("colName").Cells.Item(i + 1).Specific).Value = orset.Fields.Item("CardName").Value.ToString();
                    ((SAPbouiCOM.EditText)Matrix1.Columns.Item("colEmail").Cells.Item(i + 1).Specific).Value = orset.Fields.Item("E_Mail").Value.ToString();
                    orset.MoveFirst();
                }    

            }
        }
    }
}