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
        public static ToolStripLabel txtStats,txtUntranslated;
        public static ToolStripProgressBar PBar;
         Stack<Object[][]> undoStack = new Stack<Object[][]>();
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

                foreach (DataGridViewColumn column in Sheet.Columns)
                {
                    column.SortMode = DataGridViewColumnSortMode.NotSortable;
                }

            }
            catch (Exception)
            {
            }

        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            try
            {
                PBar.Visible = true;
                OrganizedSheet.Rows.Clear();
                OrganizedSheet.RowCount = Sheet.RowCount;
                txtStatus.Text = "Working Please wait...";
                Adapter adapter = new Adapter();
                Thread newThread = new Thread(adapter.SwitchCategory);
                newThread.Start(DropCat.SelectedIndex);
            }
            catch
            {
                MessageBox.Show("No file loaded");
                PBar.Visible = false;
            }

        }



        private void toolStripButton3_Click_1(object sender, EventArgs e)
        {
            try
            {
                Adapter adapter = new Adapter();
                Thread newThread = new Thread(adapter.SwitchBulk);
                newThread.Start(DropCat.SelectedIndex);
                //  btnSave.PerformClick();
            }catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }


        private void Form1_Activated(object sender, EventArgs e)
        {
            // Perfume.init(GridView1, OrganaizedGrid, BulkGrid);
            Sheet = GridView1;
            OrganizedSheet = OrganaizedGrid;
            BulkSheet = BulkGrid;
            txtStats = txtStatus;
            PBar = ProgressBAR;
            txtUntranslated = txtTranslatedCellCount;

            foreach (DataGridViewColumn column in OrganizedSheet.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            

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
            else if (e.Control && e.KeyCode == Keys.D)
            {
                int row = OrganizedSheet.CurrentCell.RowIndex - 1, col = OrganizedSheet.CurrentCell.ColumnIndex;
                try
                {
                    OrganizedSheet.CurrentCell.Value = OrganizedSheet[col, row].Value;
                }
                catch { }
            }
            else if (e.KeyCode == Keys.Delete)
            {
                for (int x = 0; x < OrganizedSheet.SelectedCells.Count; x++)
                {
                    OrganizedSheet.SelectedCells[x].Value = string.Empty;
                }
            }
        }
        public void pasteCell(string copied)
        {
            undoStack.Push(OrganizedSheet.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r => r.Cells.Cast<DataGridViewCell>().Select(c => c.Value).ToArray()).ToArray());

            string[] Lines = Regex.Split(copied.TrimEnd("\r\n".ToCharArray()), "\r\n");

            int StartingRow = OrganaizedGrid.CurrentCell.RowIndex;
            int StartingColumn = OrganaizedGrid.CurrentCell.ColumnIndex;
            
            foreach (var line in Lines)
            {
                if (StartingRow <= (OrganaizedGrid.Rows.Count - 1))
                {
                    string[] cells = line.Split('\t');
                    int ColumnIndex = StartingColumn;
                    for (int i = 0; i < cells.Length && ColumnIndex <= (OrganaizedGrid.Columns.Count - 1); i++)
                    {
                       
                        OrganaizedGrid[ColumnIndex++, StartingRow].Value = cells[i];

                    }


                    StartingRow++;
                }
            }
         
          
           
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {


            inputDialog dialog = new inputDialog();
            dialog.ShowDialog(this);

            dialog.Dispose();

        }

        private void Undo_Click(object sender, EventArgs e)
        {
            if (undoStack.Count == 0)
                return;
            object[][] rows = undoStack.Pop();
            while (rows.Equals(OrganizedSheet.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).ToArray()))
            {
                rows = undoStack.Pop();
            }

            OrganizedSheet.Rows.Clear();
            for (int x = 0; x <= rows.GetUpperBound(0); x++)
            {
                OrganizedSheet.Rows.Add(rows[x]);
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
            try
            {
                OrganizedSheet.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = OrganizedSheet.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue;
            }
            catch { }
        }

        private void OrganaizedGrid_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {


            undoStack.Push(OrganizedSheet.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r => r.Cells.Cast<DataGridViewCell>().Select(c => c.Value).ToArray()).ToArray());


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
                    if (OrganizedSheet.Rows[i].Cells[j].Value == null || OrganizedSheet.Rows[i].Cells[j].Value.ToString() == "" || OrganizedSheet.Rows[i].Cells[j] == null)
                    {
                        OrganizedSheet.Rows[i].Cells[j].Style.BackColor = Color.Yellow;
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


            //  Adapter adapter = new Adapter();
            //   adapter.switchDrop(DropCat.SelectedIndex, e);
        }

        
        private void toolStripButton4_Click_1(object sender, EventArgs e)
        {
            Database DB = new Database();
            //MessageBox.Show(EnglishTxtBox.Lines.Count().ToString());
            for (int x = 0; x < EnglishTxtBox.Lines.Count(); x++)
            {
                DB.AddRecord(EnglishTxtBox.Lines[x], ArabicTxtBox.Lines[x]);
            }
        }

        private void toolStripMenuItem2_Click_2(object sender, EventArgs e)
        {

        }

        private void GridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void OrganaizedGrid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {

        }

        private void OrganaizedGrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                OrganizedSheet.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = OrganizedSheet.Rows[e.RowIndex].Cells[e.ColumnIndex].FormattedValue;
            }
            catch { }
        }

        private void OrganaizedGrid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (OrganizedSheet[e.ColumnIndex, e.RowIndex].GetType() == typeof(DataGridViewComboBoxCell))
                {
                    OrganizedSheet.BeginEdit(true);
                    ComboBox comb = (ComboBox)OrganizedSheet.EditingControl;
                    comb.DroppedDown = true;
                }
            }
            catch { }
        }

        private void menu_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            //  OrganizedSheet.SelectedCells[0].Value = e.ClickedItem.Text;
        }

        private void OrganaizedGrid_CellContextMenuStripChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void OrganaizedGrid_CellContextMenuStripNeeded(object sender, DataGridViewCellContextMenuStripNeededEventArgs e)
        {


        }

        private void OrganaizedGrid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
           
          
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void doneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Sheet.SelectedCells[0].OwningRow.DefaultCellStyle.BackColor = Color.LightGreen;
        }

        private void cancelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Sheet.SelectedCells[0].OwningRow.DefaultCellStyle.BackColor = Color.HotPink;
        }

        private void doneToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            OrganizedSheet.SelectedCells[0].OwningRow.DefaultCellStyle.BackColor = Color.LightGreen;
        }

        private void cancelToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            OrganizedSheet.SelectedCells[0].OwningRow.DefaultCellStyle.BackColor = Color.HotPink;
        }

        private void removeRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            undoStack.Push(OrganizedSheet.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r => r.Cells.Cast<DataGridViewCell>().Select(c => c.Value).ToArray()).ToArray());
            try
            {
                OrganizedSheet.Rows.RemoveAt(OrganizedSheet.SelectedCells[0].RowIndex);
            }
            catch { }
        }

        private void insertColumnToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
            column.HeaderText = "Comments";
            Sheet.Columns.Insert(Sheet.SelectedCells[0].ColumnIndex , column);
        }

        private void OrganaizedGrid_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
           
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {

            PBar.Visible = true;
            txtStatus.Text = "Saving...";
            Thread thread = new Thread(() => exportToExcel());
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();


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
                    txtStats.GetCurrentParent().Invoke(new Action(() => Form1.txtStats.Text = "File has been saved!!!"));
                    PBar.GetCurrentParent().Invoke(new Action(() => Form1.PBar.Visible = false));
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
   
}




