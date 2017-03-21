using System;
using System.Windows.Forms;
using System.Configuration;
using System.IO;
using System.Text;

namespace XLStoHTMLconvert
{
    public partial class frmMain : Form
    {
        private Boolean editMode = false;

        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            tbYear.Text = ConfigurationManager.AppSettings["Year"];
            Helper.InitDictionaryAndList();       
        }

        private void btnOpenXls_Click(object sender, EventArgs e)
        {
            string fileName;
            StringBuilder week = new StringBuilder();
            
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                fileName = Path.GetFileName(openFileDialog.FileName);
                
                for (int i = 0; i < fileName.Length; i++)
                {
                    if (char.IsDigit(fileName[i]))
                    {
                        week.Append(fileName[i]);
                    }
                }

                tbWeek.Text = week.ToString();
                dateTimePicker.Value = Helper.FirstDateOfWeek(Int32.Parse(tbYear.Text), Int32.Parse(tbWeek.Text));
                dataGridView.DataSource = Helper.ImportXLS(openFileDialog.FileName);
                FormatGrid();                
            }
        }        

        private void tbWeek_Leave(object sender, EventArgs e)
        {
            try
            {
                dateTimePicker.Value = Helper.FirstDateOfWeek(Int32.Parse(tbYear.Text), Int32.Parse(tbWeek.Text));
            }
            catch (FormatException ex)
            {
                MessageBox.Show(ex.Message, "Wrong parameter format", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk);
            }                        
        }

        private void FormatGrid()
        {
            while (string.IsNullOrEmpty(dataGridView[dataGridView.Columns.Count - 1, 0].Value.ToString()))
            {
                dataGridView.Columns.RemoveAt(dataGridView.Columns.Count - 1);
            }

            while (string.IsNullOrEmpty(dataGridView[0, dataGridView.Rows.Count - 1].Value.ToString()))
            {
                dataGridView.Rows.RemoveAt(dataGridView.Rows.Count - 1);
            }

            dataGridView[0, 0].Value = tbWeek.Text + ". hét";

            for (int r = 2; r < dataGridView.Rows.Count; r++)
            {
                dataGridView[0, r].Value = Helper.FormatFirstCell(dateTimePicker.Value.AddDays(r - 2));
            }

            for (int r = 2; r < dataGridView.Rows.Count; r++)
            {
                for (int c = 1; c < dataGridView.Columns.Count; c++)
                {
                    dataGridView[c, r].Value = Helper.FormatCellData(dataGridView[c, r].Value.ToString());
                }

                dataGridView[7, r].Value = Helper.Format7thCell(dataGridView[7, r].Value.ToString(), dataGridView[1, r].Value.ToString(), dataGridView[2, r].Value.ToString());
                dataGridView[8, r].Value = Helper.Format8thCell(dataGridView[8, r].Value.ToString());                
            }
        }        

        private void dataGridView_ColumnAdded(object sender, DataGridViewColumnEventArgs e)
        {
            e.Column.SortMode = DataGridViewColumnSortMode.NotSortable;
        }

        private void dataGridView_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (editMode && (e.FormattedValue.ToString() != dataGridView[e.ColumnIndex, e.RowIndex].Value.ToString()))
            {
                Helper.SaveToDictionary(dataGridView[e.ColumnIndex, e.RowIndex].Value.ToString(), e.FormattedValue.ToString());                
            }
        }

        private void dataGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            editMode = true;
        }

        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            editMode = false;
        }

        private void btnConvertBootstrap_Click(object sender, EventArgs e)
        {
            richTextBox.Text = Helper.ConvertTableToBootstrapTable(dataGridView).ToString();
            tabControl.SelectedTab = tabHTML;
        }
    }
}
