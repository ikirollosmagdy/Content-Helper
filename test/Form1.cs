using Excel;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;




namespace helper
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }
        DataSet result;
        public static DataGridView Sheet, OrganizedSheet, BulkSheet;
        public static ToolStripLabel txtStats;
        public static ToolStripProgressBar PBar;
        Stack<changedCell> UndoAction = new Stack<changedCell>();
        Stack<changedCell[]> Undocopy = new Stack<changedCell[]>();
        public bool copiedData = false;

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            {
                OpenFileDialog OD = new OpenFileDialog();
                if (OD.ShowDialog() == DialogResult.OK)
                {
                    FileStream fs = File.Open(OD.FileName, FileMode.Open, FileAccess.Read);
                    IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                    reader.IsFirstRowAsColumnNames = true;
                    result = reader.AsDataSet();

                    ComboBox1.Items.Clear();
                    foreach (DataTable dt in result.Tables)
                        ComboBox1.Items.Add(dt.TableName);
                    reader.Close();

                }

            }
        }

        private void ComboBox1_Click(object sender, EventArgs e)
        {

        }

        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                GridView1.DataSource = result.Tables[ComboBox1.SelectedIndex];
                GridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            }
            catch (Exception)
            {
            }

        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            PBar.Visible = true;
            OrganizedSheet.Rows.Clear();
            OrganizedSheet.RowCount = Sheet.RowCount;
            txtStatus.Text = "Working Please wait...";
            Adapter adapter = new Adapter();
            Thread newThread = new Thread(adapter.SwitchCategory);
            newThread.Start(DropCat.SelectedIndex);


        }



        private void toolStripButton3_Click_1(object sender, EventArgs e)
        {
            // Creating the bulk for perfume 
            Adapter adapter = new Adapter();
            Thread newThread = new Thread(adapter.SwitchBulk);
            newThread.Start(DropCat.SelectedIndex);
          //  btnSave.PerformClick();
        }


        private void Form1_Activated(object sender, EventArgs e)
        {
            // Perfume.init(GridView1, OrganaizedGrid, BulkGrid);
            Sheet = GridView1;
            OrganizedSheet = OrganaizedGrid;
            BulkSheet = BulkGrid;
            txtStats = txtStatus;
            PBar = ProgressBAR;
           






        }

        private void OrganaizedGrid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                DataObject d = OrganaizedGrid.GetClipboardContent();
                Clipboard.SetDataObject(d);
                e.Handled = true;
            }
            else if (e.Control && e.KeyCode == Keys.V)
            {
                string CopiedContent = Clipboard.GetText();
                  pasteCell(CopiedContent);
               

            }
        }
        public void pasteCell(string copied)
        {
            string[] Lines = Regex.Split(copied.TrimEnd("\r\n".ToCharArray()), "\r\n");
           
            int StartingRow = OrganaizedGrid.CurrentCell.RowIndex;
            int StartingColumn = OrganaizedGrid.CurrentCell.ColumnIndex;
            List<changedCell> terms = new List<changedCell>();
            foreach (var line in Lines)
            {
                if (StartingRow <= (OrganaizedGrid.Rows.Count - 1))
                {
                    string[] cells = line.Split('\t');
                    int ColumnIndex = StartingColumn;
                    for (int i = 0; i < cells.Length && ColumnIndex <= (OrganaizedGrid.Columns.Count - 1); i++)
                    {
                        changedCell copiedCell = new changedCell();
                        copiedCell.column = ColumnIndex;
                        copiedCell.row = StartingRow;
                        if (OrganaizedGrid[ColumnIndex, StartingRow].Value != null)
                        {
                            copiedCell.data = OrganaizedGrid[ColumnIndex, StartingRow].Value.ToString();
                        }
                        else
                        {
                            copiedCell.data = "";
                        }
                        terms.Add(copiedCell);
                        OrganaizedGrid[ColumnIndex++, StartingRow].Value = cells[i];
                        
                    }
                       
                    
                    StartingRow++;
                }
            }
            changedCell[] arr = terms.ToArray();
            Undocopy.Push(arr);
            copiedData = true;
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {


            inputDialog dialog = new inputDialog();
            dialog.ShowDialog(this);

            dialog.Dispose();

        }

        private void Undo_Click(object sender, EventArgs e)
        {

            if (copiedData)
            {
                if (Undocopy.Count == 0)
                {
                    return;
                }
                changedCell[] arrce = new changedCell[120];
                arrce = Undocopy.Pop();
                foreach (changedCell ce in arrce)
                {
                    OrganizedSheet.Rows[ce.row].Cells[ce.column].Value = ce.data;
                }
                copiedData = false;
            }
            else
            {

                if (UndoAction.Count == 0)
                {
                    return;
                }

                changedCell cell = new changedCell();
                cell = UndoAction.Pop();
                OrganizedSheet.Rows[cell.row].Cells[cell.column].Value = cell.data;
                
            }

        }
            
        private void OrganaizedGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void GridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

        }

        private void OrganaizedGrid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            foreach (DataGridViewRow c in OrganizedSheet.Rows)
            {

                c.Cells[e.ColumnIndex].Selected = true;
            }

        }

        private void OrganaizedGrid_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            for (int x = 0; x < OrganizedSheet.ColumnCount; x++)
            {

                OrganizedSheet.Rows[e.RowIndex].Cells[x].Selected = true;
            }
        }

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {



        }

        private void toolStripDropDownButton1_Click(object sender, EventArgs e)
        {

        }

        private void OrganaizedGrid_CellLeave(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void OrganaizedGrid_CellLeave_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void OrganaizedGrid_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {




            
            changedCell cell = new changedCell();
            cell.row = e.RowIndex;
            cell.column = e.ColumnIndex;
            if (OrganizedSheet.Rows[e.RowIndex].Cells[e.ColumnIndex].Value == null)
            {
                cell.data = "";
            }
            else
            {
                cell.data = OrganizedSheet.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();
            }
            UndoAction.Push(cell);
           

        }

        private void OrganaizedGrid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnQC_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < OrganizedSheet.RowCount; i++)
            {
                for (int j = 0; j < OrganizedSheet.ColumnCount; j++)
                {
                    if (OrganizedSheet.Rows[i].Cells[j].Value == null || OrganizedSheet.Rows[i].Cells[j].Value.ToString() ==""|| OrganizedSheet.Rows[i].Cells[j]==null)
                    {
                        OrganizedSheet.Rows[i].Cells[j].Style.BackColor = Color.Red;
                    }
                }

            }
        }

        private void OrganaizedGrid_CurrentCellChanged(object sender, EventArgs e)
        {
        
        }

        private void toolStripMenuItem2_Click_1(object sender, EventArgs e)
        {
            
        }

        private void OrganaizedGrid_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            //  TextBox TB = (TextBox)e.Control;
            // TB.Multiline = true;
           // ((DataGridViewTextBoxEditingControl)e.Control).AcceptsReturn = true;
        }

        private void toolStripButton4_Click_1(object sender, EventArgs e)
        {
            Database DB = new Database();
            //MessageBox.Show(EnglishTxtBox.Lines.Count().ToString());
            for(int x = 0; x < EnglishTxtBox.Lines.Count(); x++)
            {
                DB.AddRecord(EnglishTxtBox.Lines[x], ArabicTxtBox.Lines[x]);
            }
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
                 
            /*
            Thread thread = new Thread(() => exportToExcel());
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
            */

        }
        public void exportToExcel()
        {
            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;

            try
            {

                worksheet = workbook.ActiveSheet;

                worksheet.Name = "Bulk";


                //Loop through each row and read value from each column. 
                for (int i = 0; i < BulkSheet.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1] = BulkSheet.Columns[i].HeaderText;
                }
                for (int i = 0; i < BulkSheet.Rows.Count; i++)
                {
                    for (int j = 0; j < BulkSheet.Columns.Count; j++)

                        if (BulkSheet.Rows[i].Cells[j].Value != null)
                        {
                            worksheet.Cells[i + 2, j + 1] = BulkSheet.Rows[i].Cells[j].Value.ToString();
                        }
                        else
                        {
                            worksheet.Cells[i + 2, j + 1] = "";
                        }
                }

                //Getting the location and file name of the excel to save from user. 
                SaveFileDialog saveDialog = new SaveFileDialog();
                saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                saveDialog.FilterIndex = 1;

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    MessageBox.Show("Export Successful");
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
            finally
            {
                workbook.Close(false, Type.Missing, Type.Missing);
                excel.Quit();
                workbook = null;
                excel = null;
            }

        }


    }
    class changedCell
    {

        public int row;
        public int column;
        public string data;
//        public DataGridViewCellStyle style;
    }
}




