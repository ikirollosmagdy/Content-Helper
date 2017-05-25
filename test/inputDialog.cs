using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace helper
{
    public partial class inputDialog : Form
    {
        int Count = 0;
        public inputDialog()
        {
            InitializeComponent();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            for (int x = 0; x < Form1.OrganizedSheet.SelectedCells.Count; x++)
            {
                if (Form1.OrganizedSheet.SelectedCells[x].Value != null)
                {
                    if (txtReplace.Text == string.Empty) {
                        Form1.OrganizedSheet.SelectedCells[x].Value = Form1.OrganizedSheet.SelectedCells[x].Value.ToString().ToLower().Replace(txtSearch.Text, string.Empty);
                    }
                    else
                    {
                        Form1.OrganizedSheet.SelectedCells[x].Value = Form1.OrganizedSheet.SelectedCells[x].Value.ToString().ToLower().Replace(txtSearch.Text, txtReplace.Text);
                    }
                    Count++;
                }


            }
            MessageBox.Show(string.Format("{0} cell(s) replaced", Count));
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            for (int x = Form1.OrganizedSheet.SelectedCells.Count-1; x > 0 ; x--)
            {
                if (Form1.OrganizedSheet.SelectedCells[x].Value != null)
                {
                    if (Form1.OrganizedSheet.SelectedCells[x].Value.ToString().ToLower().Contains(txtSearch.Text.ToLower()))
                    {
                        if (txtReplace.Text == string.Empty)
                        {
                            Form1.OrganizedSheet.SelectedCells[x].Value = Form1.OrganizedSheet.SelectedCells[x].Value.ToString().Replace(txtSearch.Text, string.Empty);
                        }
                        else
                        {
                            Form1.OrganizedSheet.SelectedCells[x].Value = Form1.OrganizedSheet.SelectedCells[x].Value.ToString().Replace(txtSearch.Text, txtReplace.Text);
                            break;
                        }
                    }
                }

            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
           
        }

        private void btnFind_Click(object sender, EventArgs e)
        {
            listResult.Items.Clear();
            for (int x = 0; x < Form1.OrganizedSheet.ColumnCount; x++)
            {
                for (int y = 0; y < Form1.OrganizedSheet.RowCount; y++)
                {
                    if (Form1.OrganizedSheet[x, y].Value != null)
                    {
                        if (Form1.OrganizedSheet[x, y].Value.ToString().Trim().ToLower().Contains(txtFind.Text.ToLower().Trim()))
                        {
                            ListViewItem item = new ListViewItem(new[] { Form1.OrganizedSheet[x, y].Value.ToString(), string.Format("[{0},{1}]", x, y) });
                            listResult.Items.Add(item);
                         
                        }
                    }
                }
            }
            listResult.Visible = true;
            
        }

        private void listResult_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                string text = listResult.SelectedItems[0].SubItems[1].Text;
                int[] numbers = (from Match m in Regex.Matches(text, @"\d+") select int.Parse(m.Value)).ToArray();
                  Form1.OrganizedSheet.CurrentCell = Form1.OrganizedSheet[numbers[0], numbers[1]];
               
            }
            catch { }
        }
    }
}
