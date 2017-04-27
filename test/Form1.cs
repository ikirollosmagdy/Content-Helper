using Excel;
using Microsoft.Office.Interop.Excel;
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
        public static ToolStripLabel txtStats, txtUntranslated;
        public static ToolStripProgressBar PBar;
        public static TabControl tabs;
        public static TabPage translationTab;
        Stack<Object[][]> undoStack = new Stack<Object[][]>();
        Stack<DataGridViewCellStyle[]> undoColor = new Stack<DataGridViewCellStyle[]>();
       
        public bool copiedData = false;

        private void btnLoadFile_Click(object sender, EventArgs e)
        {
           try {
                OpenFileDialog OD = new OpenFileDialog();
                if (OD.ShowDialog() == DialogResult.OK)
                {
                    FileStream fs = File.Open(OD.FileName, FileMode.Open, FileAccess.Read);
                    IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                    reader.IsFirstRowAsColumnNames = true;
                    result = reader.AsDataSet();

                    ComboBox1.Items.Clear();
                    foreach (System.Data.DataTable dt in result.Tables)
                        ComboBox1.Items.Add(dt.TableName);
                    reader.Close();

                }

            }
            catch
            {
                MessageBox.Show("Please close file first...!!");
            }
        }



        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                Sheet.Columns.Clear();
                GridView1.DataSource = result.Tables[ComboBox1.SelectedIndex];
                GridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;

                DataGridViewTextBoxColumn columnID = new DataGridViewTextBoxColumn();
                columnID.HeaderText = "No";
                if (!columnID.HeaderText.Equals(Sheet.Columns[0].HeaderText)){
                    Sheet.Columns.Insert(0, columnID);
                }
                foreach (DataGridViewRow row in Sheet.Rows)
                {
                    Sheet[0, row.Index].Value = (row.Index + 1).ToString();
                }

                foreach (DataGridViewColumn column in Sheet.Columns)
                {
                    column.SortMode = DataGridViewColumnSortMode.NotSortable;
                }

            }
            catch (Exception)
            {
            }

        }

        private void btnAnalayze_Click(object sender, EventArgs e)
        {
            try
            {
                PBar.Style = ProgressBarStyle.Marquee;
               // PBar.Visible = true;
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
                // PBar.Visible = false;
                PBar.Style = ProgressBarStyle.Continuous;
            }

        }



        private void btnCreateBulk_Click(object sender, EventArgs e)
        {
            if (OrganizedSheet.RowCount >0)
            {

                try
                {
                    txtStats.Text = "Processing";
                    PBar.Style = ProgressBarStyle.Marquee;
                    Adapter adapter = new Adapter();
                    Thread newThread = new Thread(adapter.SwitchBulk);
                    newThread.Start(DropCat.SelectedIndex);
                    tabControl1.SelectedIndex = 1;

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Please, Analayze first...!", "Warning",MessageBoxButtons.OK,MessageBoxIcon.Warning);
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
            tabs = tabControl1;
            translationTab = tabTranslation;

            foreach (DataGridViewColumn column in OrganizedSheet.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            

        }

        private void OrganaizedGrid_KeyDown(object sender, KeyEventArgs e)
        {
            KeyDownFunction(OrganizedSheet, e);
        }
        public void pasteCell(string copied)
        {
            undoStack.Push(OrganizedSheet.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r => r.Cells.Cast<DataGridViewCell>().Select(c => c.Value).ToArray()).ToArray());
            undoColor.Push(OrganizedSheet.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r => r.DefaultCellStyle).ToArray());


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

        private void ReplaceOrgMenu_Click(object sender, EventArgs e)
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
            DataGridViewCellStyle[] colors = undoColor.Pop();
            while (rows.Equals(OrganizedSheet.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).ToArray()))
            {
                rows = undoStack.Pop();
                colors = undoColor.Pop();
            }



            OrganizedSheet.Rows.Clear();
            for (int x = 0; x <= rows.GetUpperBound(0); x++)
            {
              
               OrganizedSheet.Rows.Add(rows[x]);
                OrganaizedGrid.Rows[x].DefaultCellStyle = colors[x];





            }


        }

        private void OrganaizedGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void GridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            Sheet.SelectionMode = DataGridViewSelectionMode.ColumnHeaderSelect;
            Sheet.Columns[e.ColumnIndex].Selected = true;
        }

        private void OrganaizedGrid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            OrganizedSheet.SelectionMode = DataGridViewSelectionMode.ColumnHeaderSelect;
            OrganizedSheet.Columns[e.ColumnIndex].Selected = true;


        }

        private void OrganaizedGrid_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            OrganizedSheet.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
            OrganizedSheet.Rows[e.RowIndex].Selected = true;
        }

        private void toolStripStatusLabel1_Click(object sender, EventArgs e)
        {

        }




        private void OrganaizedGrid_CellLeave(object sender, DataGridViewCellEventArgs e)
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
            undoColor.Push(OrganizedSheet.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r => r.DefaultCellStyle).ToArray());
           
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






        private void toolStripButton4_Click_1(object sender, EventArgs e)
        {
            Thread thread = new Thread(saveToDataBase);
                        thread.Start();
            MessageBox.Show("Saved");
            
            
        }
        private void saveToDataBase()
        {
            Database DB = new Database();

            for (int x = 0; x < EnglishTxtBox.Lines.Count(); x++)
            {
                DB.AddRecord(EnglishTxtBox.Lines[x], ArabicTxtBox.Lines[x]);
            }
        }


        private void GridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
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
                if (OrganizedSheet[e.ColumnIndex,e.RowIndex].Value.ToString().Contains("http"))
                {
                    ImageForm img = new ImageForm();
                    img.Url = OrganizedSheet[e.ColumnIndex,e.RowIndex].Value.ToString();
                    img.Show();
                }
                if (OrganizedSheet[e.ColumnIndex, e.RowIndex].GetType() == typeof(DataGridViewComboBoxCell))
                {
                    OrganizedSheet.BeginEdit(true);
                    ComboBox comb = (ComboBox)OrganizedSheet.EditingControl;
                    comb.DroppedDown = true;
                }
            }
            catch { }
        }



        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void doneToolMenuSheet_Click(object sender, EventArgs e)
        {
            try
            {
                for (int x = 0; x < Sheet.SelectedCells.Count; x++)
                {
                    int selectedRow = Sheet.SelectedCells[x].OwningRow.Index;
                    Sheet.Rows[selectedRow].DefaultCellStyle.BackColor = Color.LightGreen;
                    OrganizedSheet.Rows[selectedRow].DefaultCellStyle.BackColor = Color.LightGreen;
                }
            }
            catch { }
        }

        private void cancelToolSheetMenu_Click(object sender, EventArgs e)
        {
            try
            {
                for (int x = 0; x < Sheet.SelectedCells.Count; x++)
                {
                    int selectedRow = Sheet.SelectedCells[x].OwningRow.Index;
                    Sheet.Rows[selectedRow].DefaultCellStyle.BackColor = Color.HotPink;
                    OrganizedSheet.Rows[selectedRow].DefaultCellStyle.BackColor = Color.HotPink;
                }
            }
            catch { }
        }

        private void doneToolOrgMenu_Click(object sender, EventArgs e)
        {
            try
            {
                for (int x = 0; x < OrganizedSheet.SelectedCells.Count; x++)
                {
                    int selectedRow = OrganizedSheet.SelectedCells[x].OwningRow.Index;
                    OrganizedSheet.Rows[selectedRow].DefaultCellStyle.BackColor = Color.LightGreen;
                    Sheet.Rows[selectedRow].DefaultCellStyle.BackColor = Color.LightGreen;
                }
            }
            catch { }


        }

        private void cancelToolOrgMenu_Click(object sender, EventArgs e)
        {
            try
            {
                for (int x = 0; x < OrganizedSheet.SelectedCells.Count; x++)
                {
                    int selectedRow = OrganizedSheet.SelectedCells[x].OwningRow.Index;
                    OrganizedSheet.Rows[selectedRow].DefaultCellStyle.BackColor = Color.HotPink;
                    Sheet.Rows[selectedRow].DefaultCellStyle.BackColor = Color.HotPink;
                }
            }
            catch { }
        }

        private void removeRowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            undoStack.Push(OrganizedSheet.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r => r.Cells.Cast<DataGridViewCell>().Select(c => c.Value).ToArray()).ToArray());
            undoColor.Push(OrganizedSheet.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r => r.DefaultCellStyle).ToArray());
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
            Sheet.Columns.Insert(Sheet.SelectedCells[0].ColumnIndex, column);
        }

        private void btnSaveSheet_Click(object sender, EventArgs e)
        {
            PBar.Visible = true;
            txtStatus.Text = "Saving...";
            Thread thread = new Thread(() => exportToExcel(Sheet, OrganizedSheet, "Organized"));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }



        private void OrganaizedGrid_SelectionChanged_1(object sender, EventArgs e)
        {

        }

        private void OrganaizedGrid_CurrentCellChanged_1(object sender, EventArgs e)
        {

            try
            {

                Sheet.CurrentCell = Sheet[0, OrganizedSheet.CurrentCell.RowIndex];
            }
            catch { }

        }

        private void GridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult result = MessageBox.Show("Do you want to save changes?", "Confirmation", MessageBoxButtons.YesNoCancel);
            if (result == DialogResult.Yes)
            {
                try
                {
                    e.Cancel = true;
                    PBar.Visible = true;
                    txtStatus.Text = "Saving...";
                    exportToExcel(Sheet, OrganizedSheet, "Organized");
                    e.Cancel = false;

                }
                catch { }
            }
            else if (result == DialogResult.No)
            {
                e.Cancel = false;
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void toolStripButton5_Click_1(object sender, EventArgs e)
        {

            


        }

        private void CleartoolOrgMenu_Click(object sender, EventArgs e)
        {
            try
            {
                for (int x = 0; x < OrganizedSheet.SelectedCells.Count; x++)
                {
                    int selectedRow = OrganizedSheet.SelectedCells[x].OwningRow.Index;
                    OrganizedSheet.Rows[selectedRow].DefaultCellStyle.BackColor = Color.Empty;
                    Sheet.Rows[selectedRow].DefaultCellStyle.BackColor = Color.Empty;
                }
            }
            catch { }
        }

        private void CleartoolSheetMenu_Click(object sender, EventArgs e)
        {
            try
            {
                for (int x = 0; x < Sheet.SelectedCells.Count; x++)
                {
                    int selectedRow = Sheet.SelectedCells[x].OwningRow.Index;
                    Sheet.Rows[selectedRow].DefaultCellStyle.BackColor = Color.Empty;
                    OrganizedSheet.Rows[selectedRow].DefaultCellStyle.BackColor = Color.Empty;
                }
            }
            catch { }
        }

        private void GridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            Sheet.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
            Sheet.Rows[e.RowIndex].Selected = true;
        }

        private void GridView1_KeyDown(object sender, KeyEventArgs e)
        {
            KeyDownFunction(Sheet, e);
        }

        private void OrganaizedGrid_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                txtCellContent.Text = OrganizedSheet[e.ColumnIndex, e.RowIndex].Value.ToString();
                
            }
            catch {
                txtCellContent.Text = string.Empty;
            }
        }

        private void GridView1_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                txtCellContent.Text = Sheet[e.ColumnIndex, e.RowIndex].Value.ToString();

              

            }
            catch
            {
                txtCellContent.Text = string.Empty;
            }
        }

        private void BulkGrid_CellMouseMove(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                txtCellContent.Text = BulkSheet[e.ColumnIndex, e.RowIndex].Value.ToString();
            }
            catch
            {
                txtCellContent.Text = string.Empty;
            }
        }

        private void OrganaizedGrid_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            if (Properties.Settings.Default.IsFirstUse)
            {
                Splash_Screen sp = new Splash_Screen();
                sp.ShowDialog();
            }
            tabControl1.TabPages.Remove(tabTranslation);
         
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
           
        }

        private void toolStripDropDownButton1_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            MessageBox.Show(e.ClickedItem.MergeIndex.ToString());
        }

        private void Form1_HelpButtonClicked(object sender, CancelEventArgs e)
        {
          
        }

        private void btnChooseCountry_Click(object sender, EventArgs e)
        {
            Splash_Screen sp = new Splash_Screen();
            sp.ShowDialog();
        }

        private void btnTranslation_Click(object sender, EventArgs e)
        {
            TranslationUsers user = new TranslationUsers();
            user.ShowDialog();
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {

            PBar.Visible = true;
            txtStatus.Text = "Saving...";
            Thread thread = new Thread(() => exportToExcel(BulkSheet, null, "Bulk"));
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();


        }
        public void exportToExcel(DataGridView Grid, DataGridView Grid2, string SheetName)
        {
            //Getting the location and file name of the excel to save from user. 
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            saveDialog.FilterIndex = 1;
            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;
            workbook.Sheets.Add();
            try
            {

                worksheet = workbook.Sheets[1];

                worksheet.Name = SheetName;

                //Loop through each row and read value from each column. 
                for (int i = 0; i < Grid.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1] = Grid.Columns[i].HeaderText;
                }
                for (int i = 0; i < Grid.Rows.Count; i++)
                {
                    for (int j = 0; j < Grid.Columns.Count; j++)

                        if (Grid.Rows[i].Cells[j].Value != null)
                        {
                            Range range = (Range)worksheet.Cells[i + 2, j + 1];
                            worksheet.Cells[i + 2, j + 1] = Grid.Rows[i].Cells[j].Value.ToString();
                            if (Grid.Rows[i].DefaultCellStyle.BackColor == Color.Empty)
                            {
                                range.Interior.ColorIndex = 0;
                            }
                            else
                            {
                                range.Interior.Color = System.Drawing.ColorTranslator.ToOle(Grid.Rows[i].DefaultCellStyle.BackColor);
                            }


                        }
                        else
                        {
                            worksheet.Cells[i + 2, j + 1] = "";
                        }
                }
                if (Grid2 != null)
                {

                    worksheet = workbook.Sheets[2];

                    worksheet.Name = "Imported";

                    //Loop through each row and read value from each column. 
                    for (int x = 0; x < Grid2.Columns.Count; x++)
                    {
                        worksheet.Cells[1, x + 1] = Grid2.Columns[x].HeaderText;
                    }
                    for (int x = 0; x < Grid2.Rows.Count; x++)
                    {
                        for (int y = 0; y < Grid2.Columns.Count; y++)

                            if (Grid2.Rows[x].Cells[y].Value != null)
                            {
                                Range range = (Range)worksheet.Cells[x + 2, y + 1];
                                worksheet.Cells[x + 2, y + 1] = Grid2.Rows[x].Cells[y].Value.ToString();
                                if (Grid2.Rows[x].DefaultCellStyle.BackColor == Color.Empty)
                                {
                                    range.Interior.ColorIndex = 0;
                                }
                                else
                                {
                                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(Grid2.Rows[x].DefaultCellStyle.BackColor);
                                }
                            }
                            else
                            {
                                worksheet.Cells[x + 2, y + 1] = "";
                            }
                    }



                }

                if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    workbook.SaveAs(saveDialog.FileName);
                    txtStats.GetCurrentParent().Invoke(new System.Action(() => Form1.txtStats.Text = "File has been saved!!!"));
                    PBar.GetCurrentParent().Invoke(new System.Action(() => Form1.PBar.Visible = false));
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

        private void KeyDownFunction(DataGridView Grid, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.C)
            {
                DataObject d = Grid.GetClipboardContent();
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
                undoStack.Push(Grid.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r => r.Cells.Cast<DataGridViewCell>().Select(c => c.Value).ToArray()).ToArray());
                undoColor.Push(OrganizedSheet.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r => r.DefaultCellStyle).ToArray());
                try
                {
                    for (int x = Grid.SelectedCells.Count - 1; x >= 0; x--)
                    {
                        if (Grid[Grid.SelectedCells[x].ColumnIndex, Grid.SelectedCells[x].RowIndex + 1].Selected)
                        {
                            Grid[Grid.SelectedCells[x].ColumnIndex, Grid.SelectedCells[x].RowIndex + 1].Value
                                = Grid[Grid.SelectedCells[x].ColumnIndex, Grid.SelectedCells[x].RowIndex].Value;
                        }
                    }

                }
                catch { }


                try
                {

                    for (int x = Grid.SelectedCells.Count - 1; x >= 0; x--)
                    {
                        if (Grid[Grid.SelectedCells[x].ColumnIndex, Grid.SelectedCells[x].RowIndex].Selected)
                        {
                            Grid[Grid.SelectedCells[x].ColumnIndex, Grid.SelectedCells[x].RowIndex].Value
                                = Grid[Grid.SelectedCells[x].ColumnIndex, Grid.SelectedCells[0].RowIndex - 1].Value;
                        }
                    }

                }
                catch { }
            }
            else if (e.KeyCode == Keys.Delete || e.KeyCode == Keys.Back)
            {
                for (int x = 0; x < Grid.SelectedCells.Count; x++)
                {
                    Grid.SelectedCells[x].Value=DBNull.Value;
                }
            }
        }
    }

}




