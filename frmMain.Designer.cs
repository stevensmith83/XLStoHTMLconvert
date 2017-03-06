namespace XLStoHTMLconvert
{
    partial class frmMain
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
            this.btnOpenXls = new System.Windows.Forms.Button();
            this.openFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabSource = new System.Windows.Forms.TabPage();
            this.btnConvert = new System.Windows.Forms.Button();
            this.dataGridView = new System.Windows.Forms.DataGridView();
            this.dateTimePicker = new System.Windows.Forms.DateTimePicker();
            this.tbWeek = new System.Windows.Forms.TextBox();
            this.tbYear = new System.Windows.Forms.TextBox();
            this.tabHTML = new System.Windows.Forms.TabPage();
            this.richTextBox = new System.Windows.Forms.RichTextBox();
            this.bindingSource1 = new System.Windows.Forms.BindingSource(this.components);
            this.btnConvertBootstrap = new System.Windows.Forms.Button();
            this.tabControl.SuspendLayout();
            this.tabSource.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).BeginInit();
            this.tabHTML.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).BeginInit();
            this.SuspendLayout();
            // 
            // btnOpenXls
            // 
            this.btnOpenXls.Location = new System.Drawing.Point(6, 6);
            this.btnOpenXls.Name = "btnOpenXls";
            this.btnOpenXls.Size = new System.Drawing.Size(75, 23);
            this.btnOpenXls.TabIndex = 0;
            this.btnOpenXls.Text = "Open XLS";
            this.btnOpenXls.UseVisualStyleBackColor = true;
            this.btnOpenXls.Click += new System.EventHandler(this.btnOpenXls_Click);
            // 
            // openFileDialog
            // 
            this.openFileDialog.Filter = "XLS files|*.xls|All files|*.*";
            this.openFileDialog.Title = "Open XLS";
            // 
            // tabControl
            // 
            this.tabControl.Controls.Add(this.tabSource);
            this.tabControl.Controls.Add(this.tabHTML);
            this.tabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl.Location = new System.Drawing.Point(0, 0);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(1215, 635);
            this.tabControl.TabIndex = 1;
            // 
            // tabSource
            // 
            this.tabSource.Controls.Add(this.btnConvertBootstrap);
            this.tabSource.Controls.Add(this.btnConvert);
            this.tabSource.Controls.Add(this.dataGridView);
            this.tabSource.Controls.Add(this.dateTimePicker);
            this.tabSource.Controls.Add(this.tbWeek);
            this.tabSource.Controls.Add(this.tbYear);
            this.tabSource.Controls.Add(this.btnOpenXls);
            this.tabSource.Location = new System.Drawing.Point(4, 22);
            this.tabSource.Name = "tabSource";
            this.tabSource.Padding = new System.Windows.Forms.Padding(3);
            this.tabSource.Size = new System.Drawing.Size(1207, 609);
            this.tabSource.TabIndex = 0;
            this.tabSource.Text = "Source";
            this.tabSource.UseVisualStyleBackColor = true;
            // 
            // btnConvert
            // 
            this.btnConvert.Location = new System.Drawing.Point(378, 6);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(110, 23);
            this.btnConvert.TabIndex = 5;
            this.btnConvert.Text = "Convert to HTML";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // dataGridView
            // 
            this.dataGridView.AllowUserToAddRows = false;
            this.dataGridView.AllowUserToDeleteRows = false;
            this.dataGridView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(238)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dataGridView.DefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridView.Location = new System.Drawing.Point(8, 35);
            this.dataGridView.Name = "dataGridView";
            this.dataGridView.Size = new System.Drawing.Size(1191, 566);
            this.dataGridView.TabIndex = 4;
            this.dataGridView.VirtualMode = true;
            this.dataGridView.CellBeginEdit += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dataGridView_CellBeginEdit);
            this.dataGridView.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_CellEndEdit);
            this.dataGridView.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dataGridView_CellValidating);
            this.dataGridView.ColumnAdded += new System.Windows.Forms.DataGridViewColumnEventHandler(this.dataGridView_ColumnAdded);
            // 
            // dateTimePicker
            // 
            this.dateTimePicker.Location = new System.Drawing.Point(172, 8);
            this.dateTimePicker.Name = "dateTimePicker";
            this.dateTimePicker.Size = new System.Drawing.Size(200, 20);
            this.dateTimePicker.TabIndex = 3;
            // 
            // tbWeek
            // 
            this.tbWeek.Location = new System.Drawing.Point(130, 8);
            this.tbWeek.Name = "tbWeek";
            this.tbWeek.Size = new System.Drawing.Size(36, 20);
            this.tbWeek.TabIndex = 2;
            this.tbWeek.Leave += new System.EventHandler(this.tbWeek_Leave);
            // 
            // tbYear
            // 
            this.tbYear.Location = new System.Drawing.Point(88, 8);
            this.tbYear.Name = "tbYear";
            this.tbYear.Size = new System.Drawing.Size(36, 20);
            this.tbYear.TabIndex = 1;
            // 
            // tabHTML
            // 
            this.tabHTML.Controls.Add(this.richTextBox);
            this.tabHTML.Location = new System.Drawing.Point(4, 22);
            this.tabHTML.Name = "tabHTML";
            this.tabHTML.Padding = new System.Windows.Forms.Padding(3);
            this.tabHTML.Size = new System.Drawing.Size(1207, 609);
            this.tabHTML.TabIndex = 1;
            this.tabHTML.Text = "HTML code";
            this.tabHTML.UseVisualStyleBackColor = true;
            // 
            // richTextBox
            // 
            this.richTextBox.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.richTextBox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextBox.Location = new System.Drawing.Point(3, 3);
            this.richTextBox.Name = "richTextBox";
            this.richTextBox.Size = new System.Drawing.Size(1201, 603);
            this.richTextBox.TabIndex = 0;
            this.richTextBox.Text = "";
            // 
            // btnConvertBootstrap
            // 
            this.btnConvertBootstrap.Location = new System.Drawing.Point(494, 6);
            this.btnConvertBootstrap.Name = "btnConvertBootstrap";
            this.btnConvertBootstrap.Size = new System.Drawing.Size(110, 23);
            this.btnConvertBootstrap.TabIndex = 6;
            this.btnConvertBootstrap.Text = "Convert to HTML";
            this.btnConvertBootstrap.UseVisualStyleBackColor = true;
            this.btnConvertBootstrap.Click += new System.EventHandler(this.btnConvertBootstrap_Click);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1215, 635);
            this.Controls.Add(this.tabControl);
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "XLS to HTML converter";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.tabControl.ResumeLayout(false);
            this.tabSource.ResumeLayout(false);
            this.tabSource.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView)).EndInit();
            this.tabHTML.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.bindingSource1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnOpenXls;
        private System.Windows.Forms.OpenFileDialog openFileDialog;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabSource;
        private System.Windows.Forms.TabPage tabHTML;
        private System.Windows.Forms.TextBox tbYear;
        private System.Windows.Forms.DateTimePicker dateTimePicker;
        private System.Windows.Forms.TextBox tbWeek;
        private System.Windows.Forms.DataGridView dataGridView;
        private System.Windows.Forms.RichTextBox richTextBox;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Button btnConvertBootstrap;
        private System.Windows.Forms.BindingSource bindingSource1;
    }
}

