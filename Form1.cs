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
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //29/09/2021 - Send text to the clipboard so it can be pasted back
        //Not included in BulkPicRegisteration project
        private void CopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //copies the selected data to the clipboard
            if (dataGridView1.SelectedCells.Count > 0)
            {
                Clipboard.SetDataObject(dataGridView1.GetClipboardContent());
            }
        }

        //29/09/2021 - Paste the copied text back into the datagridview
        //Not included in BulkPicRegisteration project
        private void PasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //should paste the data back into the datagrid but only works outside, e.g. on the internet
            Clipboard.GetText();
        }

        //09/10/2021 - Restrict user input to the comboboxes
        //Included in EPED project
        private void btnClickMe_Click(object sender, EventArgs e)
        {
            comboBox1.Text = InputValidation.UserValidationInput(comboBox1.Text);

            //28/10/2021 - Set user input to the fields that are shown in the EPED project
            textBox1.Text = InputValidation.UserValidationInput(textBox1.Text);
            comboBox2.Text = InputValidation.UserValidationInput(comboBox2.Text);
            comboBox3.Text = InputValidation.UserValidationInput(comboBox3.Text);
            comboBox4.Text = InputValidation.UserValidationInput(comboBox4.Text);
            comboBox5.Text = InputValidation.UserValidationInput(comboBox5.Text);
        }

        //10/11/2021 - Initliaze new table with name
        private DataTable EPED = new DataTable();

        //26/10/2021 - Change date/time so it doesn't have big gaps, e.g. from '12     May     2021' to '12/May/2021'
        //Included in EPED project
        private void Form1_Load(object sender, EventArgs e)
        {
            dateTimePickerCustom.CustomFormat = "dd/MMM/yyyy";
            dateTimePicker1.CustomFormat = "dd/MMM/yyyy";

        //10/11/2021 - Add table to datagrid when project runs
        //Included in Kanban project
            EPED = new DataTable();
            EPED.Columns.Add("ID", typeof(int));
            EPED.Columns.Add("First name", typeof(string));
            EPED.Columns.Add("Surname", typeof(string));

        //04/11/2021 - Add data in grid
        //Not included in EPED project
            EPED.Rows.Add("1", "Gemma", "Thomas-Green");
            EPED.Rows.Add("2", "Sam", "Tom");
            EPED.Rows.Add("3", "Tyler", "Jenkins");
            EPED.Rows.Add("3", "James", "Micheals");

        //10/11/2021 - Set data source
        //Included in Kanban project
            dataGridView1.DataSource = EPED;
        }

        //26/10/2021 - Add ID number that auto increments each time a new row gets added
        //To be included in EPED project
        private void dataGridView1_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            dataGridView1.Rows[e.RowIndex].Cells[0].Value = (e.RowIndex + 1).ToString();
        }

        //28/10/2021 - Change panel boarder so it has the same feel as the EFFECT Photonics website
        //Not included in EPED project
        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            ControlPaint.DrawBorder(e.Graphics, panel1.ClientRectangle,
            Color.Black, 1, ButtonBorderStyle.Solid, // left
            Color.Black, 1, ButtonBorderStyle.Solid, // top
            Color.Black, 1, ButtonBorderStyle.Solid, // right
            Color.Black, 1, ButtonBorderStyle.Solid);// bottom
        }

        //02/11/2021 - Small portian of the EPED layout changes, where the button goes back to original colour when mouse is ontop
        //Included in EPED project
        private void btnClickMe_MouseMove(object sender, MouseEventArgs e)
        {
            btnClickMe.ForeColor = Color.Red;//Button font
            btnClickMe.BackColor = Color.AliceBlue;//Button
        }

        //02/11/2021 - Small portian of the EPED layout changes, where the button changes colour when mouse is away
        //Included in EPED project
        private void btnClickMe_MouseLeave(object sender, EventArgs e)
        {
            btnClickMe.ForeColor = Color.Black;//Button font
            btnClickMe.BackColor = Color.Gainsboro;//Button
        }

        //02/11/2021 - Experiment with auto increment ID field not showing in text box
        //
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];

                txtbxForename.Text = row.Cells["Forename"].Value.ToString();
                txtbxSurname.Text = row.Cells["Surname"].Value.ToString();
            }
        }

        //10/11/2021 - Open seperate form so I can practise with the Kanban project
        private void btnKanban_Click(object sender, EventArgs e)
        {
            Kanban Kanban = new Kanban();
            Kanban.Show();
        }
    }
}