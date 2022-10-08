using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Cut__copy_and_paste
{
    public partial class Kanban : Form
    {
        #region Initialize compononets
        private DataTable _table = new DataTable();
        DateTimePicker Date = new DateTimePicker();
        #endregion

        #region Constructor
        public Kanban()
        {
            InitializeComponent();
        }
        #endregion

        #region Kanban_Load
        private void Kanban_Load(object sender, EventArgs e)
        {
            //Add table
            _table = new DataTable();
            _table.Columns.Add("Part Number", typeof(string));
            _table.Columns.Add("Date", typeof(string));
            _table.Columns.Add("Quantity Issued", typeof(string));

            //Set DataGrid to table
            ScanDataGrid.DataSource = _table;

            ExportTimer.Start();
            //Set 'Date' columnn to show current date
            ScanDataGrid.Columns["Date"].DefaultCellStyle.NullValue = DateTime.Now.ToString("MM/dd/yyyy");

            Excel OpenExcel = new Excel();
            OpenExcel.OpenExcel();

            //Add DateTimePicker to DataGrid
            ScanDataGrid.Controls.Add(Date);
            Date.Visible = false;
            Date.CustomFormat = "dd/MMM/yyyy";
            Date.TextChanged += new EventHandler(DateTimePicker_TextChange);
        }
        #endregion

        #region DataGrid CellValidating
        private void ScanDataGrid_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            switch (ExportTimer.Interval)
            {
                //When timer reaches 10 minutes (600000 milliseconds) and input is detected
                case 600000 when ScanDataGrid.Rows.Count >= 0 && ScanDataGrid.Rows != null:
                    //If the 'Date' column doesn't have the date format
                    if (ScanDataGrid.Columns["Date"].DefaultCellStyle.Format != DateTime.Now.ToString("MM/dd/yyyy"))
                    {
                        MessageBox.Show("You need to input a date in the format: dd/MM/yyyy");
                    }

                    //Get method from Excel class that adds DataGrid data to Excel spreadsheet
                    Excel ExportToExcel = new Excel();
                    ExportToExcel.ExportToExcel();

                    ExportTimer.Stop();
                    //Worksheet.UsedRange.Row + Worksheet.UsedRange.Rows.Count - 1;
                    break;
                default:
                    ExportTimer.Stop();
                    MessageBox.Show("No data entered so timer has reset");
                    ExportTimer.Start();
                    break;
            }
        }
        #endregion

        #region DataGrid CellClick
        private void ScanDataGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Date.Visible = false;

            //Mouse clicked on 'Date' column, the calender is visible to easily add a date
            switch (ScanDataGrid.Columns[e.ColumnIndex].Name)
            {
                case "Date":
                    Rectangle displayCalender = ScanDataGrid.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
                    ScanDataGrid.Size = new Size(displayCalender.Width, displayCalender.Height);
                    ScanDataGrid.Location = new Point(displayCalender.X, displayCalender.Y);
                    Date.Visible = true;
                    break;
            }
        }
        #endregion

        #region DateTimePicker Text
        private void DateTimePicker_TextChange(object sender, EventArgs e)
        {
            //Set the DataGrid 'Date' column to the chosen date
            ScanDataGrid.CurrentCell.Value = Date.ToString();
        }
        #endregion

        #region FormClosing
        private void Kanban_FormClosing(object sender, FormClosingEventArgs e)
        {
            Excel FormClosing = new Excel();
            FormClosing.FormClosing();
        }
        #endregion

    }
}