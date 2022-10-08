
namespace Cut__copy_and_paste
{
    partial class Kanban
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Kanban));
            this.ScanDataGrid = new System.Windows.Forms.DataGridView();
            this.ExportTimer = new System.Windows.Forms.Timer(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.ScanDataGrid)).BeginInit();
            this.SuspendLayout();
            // 
            // ScanDataGrid
            // 
            this.ScanDataGrid.AllowUserToResizeRows = false;
            this.ScanDataGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells;
            this.ScanDataGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.NullValue = null;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.ScanDataGrid.DefaultCellStyle = dataGridViewCellStyle1;
            this.ScanDataGrid.Location = new System.Drawing.Point(-1, 1);
            this.ScanDataGrid.Name = "ScanDataGrid";
            this.ScanDataGrid.RowHeadersWidth = 51;
            this.ScanDataGrid.Size = new System.Drawing.Size(344, 392);
            this.ScanDataGrid.TabIndex = 0;
            this.ScanDataGrid.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.ScanDataGrid_CellClick);
            this.ScanDataGrid.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.ScanDataGrid_CellValidating);
            // 
            // ExportTimer
            // 
            this.ExportTimer.Interval = 600000;
            // 
            // Kanban
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(344, 455);
            this.Controls.Add(this.ScanDataGrid);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Kanban";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Kanban";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Kanban_FormClosing);
            this.Load += new System.EventHandler(this.Kanban_Load);
            ((System.ComponentModel.ISupportInitialize)(this.ScanDataGrid)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView ScanDataGrid;
        private System.Windows.Forms.Timer ExportTimer;
    }
}