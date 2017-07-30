using Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Microsoft.VisualBasic;
using System.Diagnostics;

using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using Google.Apis.Auth.OAuth2;
using System.Runtime.InteropServices;

namespace helper
{
    public partial class Form1 : Form
    {
        [DllImport("wininet.dll")]
        public extern static bool InternetGetConnectedState(out int Description, int ReservedValue);

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
        public static System.Windows.Forms.RichTextBox Englishtxt;
        public bool copiedData = false, IsEdited = false;
        Stopwatch STImported, STBulk, STTranslation;
        static string path;
        public static int LogActionsTotal;
         List<string> LogSavedFile = new List<string>();
        List<string> LogItemTypes = new List<string>();
        List<int> LogActionList = new List<int>();
        int LogTranslatedLines, LogListedItems = 0, LogActionsPerType = 0;
        bool ViewOrientation = false;

        private void btnLoadFile_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog OD = new OpenFileDialog();
                OD.Filter = "Excel files (*.xlsx)|*.xlsx";
                OD.FilterIndex = 0;
                if (OD.ShowDialog() == DialogResult.OK)
                {
                    
                    FileStream fs = File.Open(OD.FileName, FileMode.Open, FileAccess.ReadWrite);
                    IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                                     
                    reader.IsFirstRowAsColumnNames = true;
                    
                    
                    result = reader.AsDataSet();
                    
                    
                    reader.Close();
                    
                    ComboBox1.Items.Clear();
                    foreach (System.Data.DataTable dt in result.Tables)
                        ComboBox1.Items.Add(dt.TableName);
                    
                }


                try
                {
                   //Sheet.Columns.Clear();
                    ComboBox1.SelectedIndex = 0;
                    result.Tables[0].Columns[2].DataType = typeof(string);

                    GridView1.DataSource = result.Tables[0];
                    

                        GridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithAutoHeaderText;

                        DataGridViewTextBoxColumn columnID = new DataGridViewTextBoxColumn();
                        columnID.HeaderText = "No";
                        if (!columnID.HeaderText.Equals(Sheet.Columns[0].HeaderText))
                        {
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

                        LogWrite("Sheet imported \"" + ComboBox1.SelectedItem.ToString() + "\"");
                    }
                    catch (Exception)
            {
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
                //Sheet.Columns.Clear();
               
                GridView1.DataSource = result.Tables[ComboBox1.SelectedIndex];
                GridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;

                DataGridViewTextBoxColumn columnID = new DataGridViewTextBoxColumn();
                columnID.HeaderText = "No";
                if (!columnID.HeaderText.Equals(Sheet.Columns[0].HeaderText))
                {
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

                LogWrite("Sheet imported \"" + ComboBox1.SelectedItem.ToString() + "\"");
            }
            catch (Exception)
            {
            }

        }

        private void btnAnalayze_Click(object sender, EventArgs e)
        {
            try
            {
                LogWrite("Start Analayzing category \"" + DropCat.SelectedItem.ToString() + "\"");
                IsEdited = true;
                PBar.Style = ProgressBarStyle.Marquee;
                OrganizedSheet.Rows.Clear();
                OrganizedSheet.RowCount = Sheet.RowCount;
                txtStatus.Text = "Working Please wait...";
                WorkerAnalyze.RunWorkerAsync(DropCat.SelectedItem);
            }
            catch
            {
                MessageBox.Show("No file loaded");
                PBar.Style = ProgressBarStyle.Continuous;
            }

        }



        private void btnCreateBulk_Click(object sender, EventArgs e)
        {
            LogWrite(string.Format("Start creating bulk for category {0}", DropCat.SelectedItem.ToString()));
            if (OrganizedSheet.RowCount > 0)
            {

                try
                {
                    EnglishTxtBox.Clear();
                    ArabicTxtBox.Clear();
                    txtStats.Text = "Processing";
                    PBar.Style = ProgressBarStyle.Marquee;
                    WorkerBulk.RunWorkerAsync(DropCat.SelectedItem);

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Please, Analayze first...!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
            Englishtxt = EnglishTxtBox;

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

            LogActionsTotal++;

        }

        private void ReplaceOrgMenu_Click(object sender, EventArgs e)
        {


            inputDialog dialog = new inputDialog();
            dialog.Show(this);



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

            LogActionsTotal++;
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

        }

        private void OrganaizedGrid_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {


           undoStack.Push(OrganizedSheet.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r => r.Cells.Cast<DataGridViewCell>().Select(c => c.Value).ToArray()).ToArray());
           undoColor.Push(OrganizedSheet.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r => r.DefaultCellStyle).ToArray());

        }

        private void OrganaizedGrid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            string value = OrganizedSheet[e.ColumnIndex, e.RowIndex].Value.ToString();
            LogWrite(string.Format("Cell [{0},{1}] value changed to \"{2}\"", e.ColumnIndex, e.RowIndex, value));
           
        }

        private void btnQC_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < OrganizedSheet.RowCount; i++)
            {
                for (int j = 0; j < OrganizedSheet.ColumnCount; j++)
                {
                    if (OrganizedSheet.Rows[i].Cells[j].Value == null || OrganizedSheet.Rows[i].Cells[j].Value.ToString() == "" || OrganizedSheet.Rows[i].Cells[j] == null)
                    {
                        OrganizedSheet.Rows[i].Cells[j].Style.BackColor = System.Drawing.Color.Yellow;
                    }
                    else
                    {
                        if (OrganizedSheet.Columns[j].HeaderText == "Price" || OrganizedSheet.Columns[j].HeaderText == "Quntity")
                        {

                            Regex Number = new Regex(@"[^\d]");
                            Match NuMatch = Number.Match(OrganizedSheet[j, i].Value.ToString());
                            if (NuMatch.Success)
                            {
                                OrganizedSheet.Rows[i].Cells[j].Style.BackColor = System.Drawing.Color.Yellow;


                            }
                            

                        }
                        else
                        {
                            OrganizedSheet.Rows[i].Cells[j].Style.BackColor = System.Drawing.Color.Empty;
                        }
                    }
                }
            }
        }






        private void toolStripButton4_Click_1(object sender, EventArgs e)
        {

            EnglishTxtBox.Text = EnglishTxtBox.Text.Trim();
            ArabicTxtBox.Text = ArabicTxtBox.Text.Trim();


            string[][] array = new string[ArabicTxtBox.Lines.Count()][];
            for (int x = 0; x < ArabicTxtBox.Lines.Count(); x++)
            {
                array[x] = new string[] { EnglishTxtBox.Lines[x], ArabicTxtBox.Lines[x] };
            }

            Workertranslation.RunWorkerAsync(array);

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
                if (OrganizedSheet[e.ColumnIndex, e.RowIndex].Value.ToString().Contains("http"))
                {
                    ImageForm img = new ImageForm(OrganizedSheet[e.ColumnIndex, e.RowIndex].Value.ToString());
                  
                    img.Show();
                }
                if (OrganizedSheet[e.ColumnIndex, e.RowIndex].GetType() == typeof(DataGridViewComboBoxCell))
                {
                    OrganizedSheet.BeginEdit(true);
                    ComboBox comb = (ComboBox)OrganizedSheet.EditingControl;
                    comb.DroppedDown = true;
                    comb.SelectionChangeCommitted += Comb_SelectionChangeCommitted;
                    comb.FormatStringChanged += Comb_SelectionChangeCommitted;

                }
            }
            catch { }
        }


        private void Comb_SelectionChangeCommitted(object sender, EventArgs e)
        {
            try
            {
                OrganaizedGrid.EndEdit();
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
                    Sheet.Rows[selectedRow].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                    OrganizedSheet.Rows[selectedRow].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
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
                    Sheet.Rows[selectedRow].DefaultCellStyle.BackColor = System.Drawing.Color.HotPink;
                    OrganizedSheet.Rows[selectedRow].DefaultCellStyle.BackColor = System.Drawing.Color.HotPink;
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
                    OrganizedSheet.Rows[selectedRow].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                    Sheet.Rows[selectedRow].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
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
                    OrganizedSheet.Rows[selectedRow].DefaultCellStyle.BackColor = System.Drawing.Color.HotPink;
                    Sheet.Rows[selectedRow].DefaultCellStyle.BackColor = System.Drawing.Color.HotPink;
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


                foreach (DataGridViewCell oneCell in OrganizedSheet.SelectedCells)
                {
                    if (oneCell.Selected)
                        OrganizedSheet.Rows.RemoveAt(oneCell.RowIndex);
                }


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
            Thread thread = new Thread(() => exportToExcel(Sheet, OrganizedSheet, "Imported"));
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

                Sheet.CurrentCell = Sheet[OrganizedSheet.CurrentCell.ColumnIndex, OrganizedSheet.CurrentCell.RowIndex];
            }
            catch { }

        }

        private void GridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (IsEdited)
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
                        Report();
                        e.Cancel = false;

                    }
                    catch { }
                }
                else if (result == DialogResult.No)
                {
                    Report();
                    e.Cancel = false;
                }
                else
                {
                    e.Cancel = true;
                }
            }
        }


        private void CleartoolOrgMenu_Click(object sender, EventArgs e)
        {
            try
            {
                for (int x = 0; x < OrganizedSheet.SelectedCells.Count; x++)
                {
                    int selectedRow = OrganizedSheet.SelectedCells[x].OwningRow.Index;
                    OrganizedSheet.Rows[selectedRow].DefaultCellStyle.BackColor = System.Drawing.Color.Empty;
                    Sheet.Rows[selectedRow].DefaultCellStyle.BackColor = System.Drawing.Color.Empty;
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
                    Sheet.Rows[selectedRow].DefaultCellStyle.BackColor = System.Drawing.Color.Empty;
                    OrganizedSheet.Rows[selectedRow].DefaultCellStyle.BackColor = System.Drawing.Color.Empty;
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
                if (e.Button == System.Windows.Forms.MouseButtons.Left)
                {
                    OrganaizedGrid[e.ColumnIndex, e.RowIndex].Selected = true;
                }
            }
            catch
            {
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


            if (IsConnectedToInternet())
            {
               
                XDocument document = XDocument.Load("http://souqforms.atwebpages.com/UpdateInfo.xml");
                var elements = document.Element("AppName");
                Version onlineVersion = new Version(elements.Element("version").Value);
                Version LocalVersion = Assembly.GetExecutingAssembly().GetName().Version;
                if (LocalVersion.CompareTo(onlineVersion) < 0)
                {
                    Updater updater = new Updater("http://souqforms.atwebpages.com/");
                    updater.ShowDialog();
                }
               
                Task task1 = new Task(() => {
                    Common_Use common = new Common_Use();
                    common.sendOfflineReport();
                    
                });
                task1.Start();

            }


            System.Windows.Forms.Timer t = new System.Windows.Forms.Timer();
            _start = DateTime.Now;
            t.Start();
            t.Tick += T_Tick;
            STImported = new Stopwatch();
            STBulk = new Stopwatch();
            STTranslation = new Stopwatch();
            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\log"))
            {
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\\log");
            }
            string dt = DateTime.Now.ToString("dd-MM-yy_HHmm");
            path = System.Windows.Forms.Application.StartupPath + "\\log\\Log_" + dt + ".txt";
            LogWrite("Initilaizing");
            STImported.Start();
        }

        public static void LogWrite(string value)
        {
            try
            {
                string nowtime = DateTime.Now.ToString("dd-MM-yy HH:mm:ss");

                TimeSpan duration = DateTime.Now - _start;
                string du = string.Format("{0:hh\\:mm\\:ss}", duration);
                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine(value + " (" + du + ")  " + "  at: " + nowtime);
                    sw.Close();

                }
              
             
            }
            catch { }

        }



        static DateTime _start;
        private void T_Tick(object sender, EventArgs e)
        {
            TimeSpan duration = DateTime.Now - _start;
            //this.Text = string.Format("Katana {0:hh\\:mm\\:ss}", duration);
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

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

        private void btnDevelop_Click(object sender, EventArgs e)
        {

            if (Interaction.InputBox("Enter your password", "Developer Section") == "admin")
            {


                DeveloperForm DForm = new DeveloperForm();
                DForm.ShowDialog();
            }
            else
            {
                MessageBox.Show("Wrong Password", "Developer password", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void OrganaizedGrid_RowsDefaultCellStyleChanged(object sender, EventArgs e)
        {


        }

        private void OrganaizedGrid_RowDefaultCellStyleChanged(object sender, DataGridViewRowEventArgs e)
        {

        }

        private void OrganaizedGrid_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            LogWrite(string.Format("Row #{0} has been deleted", e.RowIndex));
        }

        private void OrganaizedGrid_RowsDefaultCellStyleChanged_1(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (System.IO.File.Exists("katana.ex_"))
            {
                System.IO.File.Delete("katana.ex_");
            }
        }

        private void WorkerAnalyze_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            txtStatus.Text = "Analayzing Finished";
            PBar.Style = ProgressBarStyle.Continuous;
        }

        private void WorkerAnalyze_DoWork(object sender, DoWorkEventArgs e)
        {

            Adapter adapter = new Adapter();
            adapter.SwitchCategory(e.Argument);

        }

        private void WorkerBulk_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            txtStatus.Text = "Finished creating Bulk";
            PBar.Style = ProgressBarStyle.Continuous;
            if (EnglishTxtBox.Text == string.Empty)
            {
                tabControl1.SelectedIndex = 1;
            }
            else
            {
                tabControl1.SelectedIndex = 2;
            }
        }

        private void WorkerBulk_DoWork(object sender, DoWorkEventArgs e)
        {
            Adapter adapter = new Adapter();
            adapter.SwitchBulk(e.Argument);
        }


        private void ArabicTxtBox_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                Englishtxt.SelectionBackColor = System.Drawing.Color.White;
                EnglishTxtBox.SelectionFont = new System.Drawing.Font(EnglishTxtBox.SelectionFont, FontStyle.Regular);
                int indexAr = ArabicTxtBox.GetLineFromCharIndex(ArabicTxtBox.SelectionStart);
                string currentlinetext = EnglishTxtBox.Lines[indexAr];
                int fy = EnglishTxtBox.GetFirstCharIndexFromLine(indexAr);
                EnglishTxtBox.Select(fy, currentlinetext.Length);
                EnglishTxtBox.SelectionBackColor = System.Drawing.Color.LightGreen;
                EnglishTxtBox.SelectionFont = new System.Drawing.Font(EnglishTxtBox.SelectionFont, FontStyle.Bold);
                EnglishTxtBox.ScrollToCaret();
             //  ArabicTxtBox.ScrollToCaret();
            }
            catch { }
        }

        private void Workertranslation_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show(string.Format("Saved {0} line(s)",EnglishTxtBox.Lines.Count()));
            LogTranslatedLines = EnglishTxtBox.Lines.Count();
            EnglishTxtBox.Text = string.Empty;
            ArabicTxtBox.Text = string.Empty;
        }

        private void Workertranslation_DoWork(object sender, DoWorkEventArgs e)
        {
            Database db = new Database();
            string[][] array = (string[][])e.Argument;

            for (int x = 0; x < array.Length; x++)
            {
                db.AddRecord(array[x][0], array[x][1]);
                Workertranslation.ReportProgress((x/array.Length)*100);
            }




        }

        private void ComboBox1_Click(object sender, EventArgs e)
        {

        }

        private void OrganizedMenu_Opening(object sender, CancelEventArgs e)
        {

        }

        private void tabControl1_Selected(object sender, TabControlEventArgs e)
        {
            int selectedPageIndex = tabControl1.SelectedIndex;
            switch (selectedPageIndex)
            {
                case 0:
                    STImported.Start();
                    STBulk.Stop();
                    STTranslation.Stop();
                    break;
                case 1:
                    STImported.Stop();
                    STBulk.Start();
                    STTranslation.Stop();
                    break;
                case 2:
                    STImported.Stop();
                    STBulk.Stop();
                    STTranslation.Start();
                    break;
            }

        }

        private void EnglishTxtBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.A)
            {
                EnglishTxtBox.SelectAll();
            }
        }

        private void CopyStripMenuItem1_Click(object sender, EventArgs e)
        {
            DataObject d = OrganaizedGrid.GetClipboardContent();

            Clipboard.SetDataObject(d);
            
        }

        private void OrganaizedGrid_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                LogActionsPerType++;
                LogActionsTotal++;
               
            }
            catch { }
        }

        private void Workertranslation_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            PBar.Value = e.ProgressPercentage;
        }

        private void GridView1_ColumnHeaderMouseDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            //string header = Sheet.Columns[e.ColumnIndex].HeaderText;
            //Sheet.Columns[e.ColumnIndex].HeaderText = Interaction.InputBox("Change column header from " + header + " to be ?", "Rename Header Title");

        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void OrganaizedGrid_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            OrganizedSheet[e.ColumnIndex, e.RowIndex].Value = DBNull.Value;
            //MessageBox.Show("Please choose one from droplist");
            if (e.Exception!=null && e.Context== DataGridViewDataErrorContexts.Display)
            {
                //MessageBox.Show("Please choose one from droplist");
                e.Cancel = true;
                e.ThrowException = false;
               
            }
        }

        private void btnAddItem_Click(object sender, EventArgs e)
        {
            if (ViewOrientation)
            {
                splitContainer2.Orientation = Orientation.Vertical;
            }
            else {
                splitContainer2.Orientation = Orientation.Horizontal;
            }
            ViewOrientation = !ViewOrientation;
        }

        private void btnTranslateBing_Click(object sender, EventArgs e)
        {
           
         getTranslation();

        }

        private async void getTranslation() {
            Common_Use com = new Common_Use();
            ArabicTxtBox.Clear();
            for (int i = 0; i < EnglishTxtBox.Lines.Count(); i++)
            {
                ArabicTxtBox.Text+= await com.Translate(EnglishTxtBox.Lines[i]) + Environment.NewLine;
            }


        }

        private void OrganaizedGrid_DragDrop(object sender, DragEventArgs e)
        {
            string cellvalue = e.Data.GetData(DataFormats.Text).ToString();
            System.Drawing.Point cursorLocation = OrganaizedGrid.PointToClient(new System.Drawing.Point(e.X, e.Y));

            System.Windows.Forms.DataGridView.HitTestInfo hittest = OrganaizedGrid.HitTest(cursorLocation.X, cursorLocation.Y);
            if (hittest.ColumnIndex != -1
                && hittest.RowIndex != -1)
                OrganaizedGrid[hittest.ColumnIndex, hittest.RowIndex].Value = cellvalue.Trim();
           
        }

        private void OrganaizedGrid_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Copy;
        }

        private void OrganaizedGrid_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            //Draw only grid content cells not ColumnHeader cells nor RowHeader cells
            if (e.ColumnIndex > -1 & e.RowIndex > -1)
            {
                //Pen for left and top borders
                using (var backGroundPen = new Pen(e.CellStyle.BackColor, 1))
                //Pen for bottom and right borders
                using (var gridlinePen = new Pen(OrganaizedGrid.GridColor, 1))
                //Pen for selected cell borders
                using (var selectedPen = new Pen(System.Drawing.Color.Red, 1))
                {
                    var topLeftPoint = new System.Drawing.Point(e.CellBounds.Left, e.CellBounds.Top);
                    var topRightPoint = new System.Drawing.Point(e.CellBounds.Right - 1, e.CellBounds.Top);
                    var bottomRightPoint = new System.Drawing.Point(e.CellBounds.Right - 1, e.CellBounds.Bottom - 1);
                    var bottomleftPoint = new System.Drawing.Point(e.CellBounds.Left, e.CellBounds.Bottom - 1);

                    //Draw selected cells here
                    if (this.OrganaizedGrid[e.ColumnIndex, e.RowIndex].Selected)
                    {
                        //Paint all parts except borders.
                        e.Paint(e.ClipBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.Border);

                        //Draw selected cells border here
                        e.Graphics.DrawRectangle(selectedPen, new System.Drawing.Rectangle(e.CellBounds.Left, e.CellBounds.Top, e.CellBounds.Width - 1, e.CellBounds.Height - 1));

                        //Handled painting for this cell, Stop default rendering.
                        e.Handled = true;
                    }
                    //Draw non-selected cells here
                    else
                    {
                        //Paint all parts except borders.
                        e.Paint(e.ClipBounds, DataGridViewPaintParts.All & ~DataGridViewPaintParts.Border);

                        //Top border of first row cells should be in background color
                        if (e.RowIndex == 0)
                            e.Graphics.DrawLine(backGroundPen, topLeftPoint, topRightPoint);

                        //Left border of first column cells should be in background color
                        if (e.ColumnIndex == 0)
                            e.Graphics.DrawLine(backGroundPen, topLeftPoint, bottomleftPoint);

                        //Bottom border of last row cells should be in gridLine color
                        if (e.RowIndex == OrganaizedGrid.RowCount - 1)
                            e.Graphics.DrawLine(gridlinePen, bottomRightPoint, bottomleftPoint);
                        else  //Bottom border of non-last row cells should be in background color
                            e.Graphics.DrawLine(backGroundPen, bottomRightPoint, bottomleftPoint);

                        //Right border of last column cells should be in gridLine color
                        if (e.ColumnIndex == OrganaizedGrid.ColumnCount - 1)
                            e.Graphics.DrawLine(gridlinePen, bottomRightPoint, topRightPoint);
                        else //Right border of non-last column cells should be in background color
                            e.Graphics.DrawLine(backGroundPen, bottomRightPoint, topRightPoint);

                        //Top border of non-first row cells should be in gridLine color, and they should be drawn here after right border
                        if (e.RowIndex > 0)
                            e.Graphics.DrawLine(gridlinePen, topLeftPoint, topRightPoint);

                        //Left border of non-first column cells should be in gridLine color, and they should be drawn here after bottom border
                        if (e.ColumnIndex > 0)
                            e.Graphics.DrawLine(gridlinePen, topLeftPoint, bottomleftPoint);

                        //We handled painting for this cell, Stop default rendering.
                        e.Handled = true;
                    }
                }
            }
        }

        private void OrganaizedGrid_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
       
         
        }

        private void GridView1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
           
        }

        private void PasteStripMenuItem1_Click(object sender, EventArgs e)
        {
            string CopiedContent = Clipboard.GetText();
            pasteCell(CopiedContent);
        }

        private void CutStripMenuItem1_Click(object sender, EventArgs e)
        {
            undoStack.Push(OrganizedSheet.Rows.Cast<DataGridViewRow>().Where(r => !r.IsNewRow).Select(r => r.Cells.Cast<DataGridViewCell>().Select(c => c.Value).ToArray()).ToArray());
            DataObject d = OrganaizedGrid.GetClipboardContent();

            Clipboard.SetDataObject(d);
          
            foreach (DataGridViewCell cell in OrganaizedGrid.SelectedCells)
            {
                cell.Value = DBNull.Value;
            }
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
            DropCat.GetCurrentParent().Invoke(new System.Action(() => LogItemTypes.Add(DropCat.SelectedItem.ToString())));
            for (int g = 0; g < BulkSheet.RowCount; g++)
            {
                if (BulkSheet[0, g] != null || BulkSheet[0, g].Value != null || BulkSheet[0, g].Value.ToString() != string.Empty)
                {
                    LogListedItems++;
                }

            }
            //Getting the location and file name of the excel to save from user. 
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            saveDialog.FilterIndex = 1;
            if (saveDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
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
                                if (Grid.Rows[i].DefaultCellStyle.BackColor == System.Drawing.Color.Empty)
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

                        worksheet.Name = "Organized";

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
                                    if (Grid2.Rows[x].DefaultCellStyle.BackColor == System.Drawing.Color.Empty || Grid2.Rows[x].DefaultCellStyle.BackColor == System.Drawing.Color.White)
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


                    workbook.SaveAs(saveDialog.FileName);
                    txtStats.GetCurrentParent().Invoke(new System.Action(() => MessageBox.Show("File has been saved...!!!","Saving...")));
                    txtStats.GetCurrentParent().Invoke(new System.Action(() => Form1.txtStats.Text = "File has been saved!!!"));
                    PBar.GetCurrentParent().Invoke(new System.Action(() => Form1.PBar.Visible = false));


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
            LogWrite(string.Format("Finished exporting sheet with name {0}", saveDialog.FileName));

            //if (saveDialog.ShowDialog() != DialogResult.OK) {
            //    saveDialog.FileName = "N/A";
            //}
            try
            {
                LogSavedFile.Add(saveDialog.FileName);
                LogActionList.Add(LogActionsPerType);

                LogActionsTotal += LogActionsTotal;

            }
            catch
            {

            }
        }


        private void Report()
        {
            try
            {
                TimeSpan duration = DateTime.Now - _start;
                string du = string.Format("{0:mm\\:ss}", duration);
                LogWrite("\r\n\r\n+====================================================================================+");
                LogWrite(string.Format("|Total item listed: {0} items(s)\t\t\t\t\t\t\t\t\t|", LogListedItems));
                LogWrite(string.Format("|Total lines translated: {0} line(s)\t\t\t\t\t\t\t\t\t|", LogTranslatedLines));
                LogWrite(string.Format("|Start time : {0}\t\t\t\t\t\t\t\t\t|", _start));
                LogWrite(string.Format("|Finish time : {0}\t\t\t\t\t\t\t\t\t|", DateTime.Now));
                LogWrite(string.Format("|Total time elapsed: {0} Min(s):Sec(s)\t\t\t\t\t\t\t\t\t|", du));
                LogWrite(string.Format("|Active sheet elapsed time\t\t\t\t\t\t\t\t\t|"));
                LogWrite(string.Format("|   -Tab Imported: {0}\t\t\t\t\t\t\t\t\t|", STImported.Elapsed));
                LogWrite(string.Format("|   -Tab Bulk: {0}\t\t\t\t\t\t\t\t\t|", STBulk.Elapsed));
                LogWrite(string.Format("|   -Tab Translation: {0}\t\t\t\t\t\t\t\t\t|", STTranslation.Elapsed));
                LogWrite(string.Format("|Total actions by Agent : {0} Action(s)\t\t\t\t|", LogActionsTotal));
                LogWrite(string.Format("|Saved file names\t\t\t\t\t\t\t\t\t|"));
                for (int x = 0; x < LogSavedFile.Count; x++)
                {
                    LogWrite(string.Format("|   -Name:{0}\t\t\t\t\n|", LogSavedFile[x]));
                }
                LogWrite(string.Format("|Item type names\t\t\t\t\t\t\t\t\t|"));
                for (int y = 0; y < LogItemTypes.Count; y++)
                {
                    LogWrite(string.Format("|   -Item type name:{0} with {1} actions\t\t\t\t\t\n|", LogItemTypes[y], LogActionList[y]));
                }
                LogWrite("+====================================================================================+");
            }
            catch
            {
                LogWrite("Application closed by force");
            }

            if (IsConnectedToInternet())
            {
                // Posting data to google Spreadsheet 
                try
                {
                    string[] Scopes = { SheetsService.Scope.Spreadsheets };
                    string ApplicationName = "Google Sheets API .NET Quickstart";

                    UserCredential credential;

                    using (var stream =
                        new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
                    {
                        string credPath = System.Environment.GetFolderPath(
                            System.Environment.SpecialFolder.Personal);


                        credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                            GoogleClientSecrets.Load(stream).Secrets,
                            Scopes,
                            "user",
                            CancellationToken.None,
                            new FileDataStore(credPath, true)).Result;
                        Console.WriteLine("Credential file saved to: " + credPath);
                    }

                    // Create Google Sheets API service.
                    var service = new SheetsService(new BaseClientService.Initializer()
                    {
                        HttpClientInitializer = credential,
                        ApplicationName = ApplicationName,
                    });



                    String spreadsheetId2 = "1AxuO473BEWpVqOtR0wfOOPPVrhUC1ta_Xxj9i79seXA";
                    String range2 = "Sheet1!A:Z";

                    ValueRange valueRange = new ValueRange();
                    valueRange.MajorDimension = "ROWS";//"ROWS";//COLUMNS
                    string names = string.Empty;
                    string types = string.Empty;
                    for (int x = 0; x < LogSavedFile.Count; x++)
                    {
                        names += string.Format("Name: {0}\n", LogSavedFile[x]);
                    }
                    for (int y = 0; y < LogItemTypes.Count; y++)
                    {
                        types += string.Format("{0} with {1} actions\n", LogItemTypes[y], LogActionList[y]);
                    }

                    TimeSpan duration = DateTime.Now - _start;
                   
                    string du = string.Format("{0:hh\\:mm\\:ss}", duration);
                    var oblist = new List<object>() {
                    Environment.UserName, LogListedItems,
                    LogTranslatedLines,
                    _start.ToString(),
                    DateTime.Now.ToString(),
                    du,             
            string.Format("{0:hh\\:mm\\:ss}",STImported.Elapsed).Trim(),
                            string.Format("{0:hh\\:mm\\:ss}",STBulk.Elapsed).Trim(),
                      string.Format("{0:hh\\:mm\\:ss}",STTranslation.Elapsed).Trim(),
                    LogActionsTotal,
                    names.Trim(),types.Trim()
                };
                    valueRange.Values = new List<IList<object>> { oblist };

                    // SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId2, range2);
                    SpreadsheetsResource.ValuesResource.AppendRequest update = service.Spreadsheets.Values.Append(valueRange, spreadsheetId2, range2);
                    update.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                    AppendValuesResponse result2 = update.Execute();


                    // End of posting and Close app
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            else {
                string names = string.Empty;
                string types = string.Empty;
                for (int x = 0; x < LogSavedFile.Count; x++)
                {
                    names += string.Format("Name: {0}\n", LogSavedFile[x]).Trim();
                }
                for (int y = 0; y < LogItemTypes.Count; y++)
                {
                    types += string.Format("{0} with {1} actions\n", LogItemTypes[y], LogActionList[y]).Trim();
                }

                TimeSpan duration = DateTime.Now - _start;
                string du = string.Format("{0:hh\\:mm\\:ss}", duration);

                OfflineReport(Environment.UserName, LogListedItems,
                    LogTranslatedLines,
                    _start.ToString(),
                    DateTime.Now.ToString(),
                    du,
                    string.Format("{0:hh\\:mm\\:ss}", STImported.Elapsed).Trim(),
                     string.Format("{0:hh\\:mm\\:ss}", STBulk.Elapsed).Trim(),
                      string.Format("{0:hh\\:mm\\:ss}", STTranslation.Elapsed).Trim(),
                    LogActionsTotal,
                    names, types);

            }

        }

        public static bool IsConnectedToInternet()
        {
            bool returnValue = false;
            try
            {

                int Desc;
                returnValue = InternetGetConnectedState(out Desc, 0);
            }
            catch
            {
                returnValue = false;
            }
            return returnValue;
        }


        private async void OfflineReport(string Agent, int ListedItems, int Translated, string stTime, string fTime, string toTime, string tiTime, string tbTime, string ttTime, int actions, string saved, string types)
        {
            try
            {
                using (StreamWriter sw = File.AppendText("OfflineReport.dat"))
                {

                    await sw.WriteLineAsync(string.Format("{0}\t{1}\t{2}\t{3}\t{4}\t{5}\t{6}\t{7}\t{8}\t{9}\t{10}\t{11}", Agent, ListedItems, Translated, stTime, fTime, toTime, tiTime, tbTime, ttTime, actions, saved, types));
                    sw.Close();

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error for offline: {0}", ex.Message);
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
                    Grid.SelectedCells[x].Value = DBNull.Value;
                }
            }
        }
    }

}




