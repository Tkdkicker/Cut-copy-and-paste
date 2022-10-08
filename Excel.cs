using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using _Application = Microsoft.Office.Interop.Excel._Application;
using Microsoft.Office.Interop.Outlook;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace Cut__copy_and_paste
{
    class Excel
    {
        #region Initialize components
        _Application excelApp = new _Excel.Application();//Reference to application
        Workbook wb;//File opening
        Worksheet ws;//All sheets
        Sheets sheets;//One sheet
        #endregion

        #region Method 1: OpenExcel();
        public void OpenExcel()
        {
            //Open Excel file
            wb = excelApp.Workbooks.Open(@"C:\Users\GemmaThomas-Green\OneDrive - EFFECT Photonics\Kanban project\Kanban Sheet NEW testing.xlsx");

            //Able to see the Excel workbook sheet 3
            excelApp.Visible = true;
            sheets = wb.Worksheets;
            ws = sheets.get_Item(3);
        }
        #endregion

        #region Method 2: ExportToExcel();
        public void ExportToExcel()
        {
            //Add data to last row of Excel workbook sheet 3
            //ws = wb.Worksheets["Material Issued"].UsedRange.Specialcells(11).ws.Columns["A:C"].InsertData(ScanDataGrid);
        }
        #endregion

        #region Method: Email();
        public void Email()
        {
            //If the 'Current Stock' is less than or equal to the 'Re-order Point' an email gets sent
            if (ws == wb.Worksheets["Kanban Sheet"].Columns["X"] <= ws == wb.Worksheets["Kanban Sheet"].Columns["U"])
            {
                //Initialize Outlook application
                Outlook.Application outlookApp = new Outlook.Application();
                MailItem mailItem = (MailItem)outlookApp.CreateItem(OlItemType.olMailItem);

                //Set the subject
                mailItem.Subject = "New order";

                //Who to send the email to
                mailItem.To = "gemmathomasgreen@effectphotonics.com";

                //The 'Description' of the stock is equal to the 'KB Number' because this is scanned in and restocked by the amount needed ('Re-order Point' + 'Current Stock')
                mailItem.HTMLBody =
                    "This item: " + ws == wb.Worksheets["Kanban Sheet"].Columns["A"] == ws == wb.Worksheets["Kanban Sheet"].Columns["D"]
                    + " needs to be restocked by at least this amount: " + ws == wb.Worksheets["Kanban Sheet"].Columns["U"];

                //How important the email is
                mailItem.Importance = OlImportance.olImportanceHigh;

                //Send the email
                mailItem.Send();
                MessageBox.Show("Email sent");
            }
            //If 'Current Stock' is more then don't send email
            return;
        }
        #endregion

        #region Method: FormClosing();
        public void FormClosing()
        {
            DialogResult result = MessageBox.Show("There may be data that could be lost if you close", "Are you sure you want to exit?", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                //Close Excel file
                wb.Close();
                //Close Excel application
                excelApp.Quit();
            }
        }
        #endregion
    }
}