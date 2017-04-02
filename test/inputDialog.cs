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

namespace test
{
    public partial class inputDialog : Form
    {
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
                    Form1.OrganizedSheet.SelectedCells[x].Value = Form1.OrganizedSheet.SelectedCells[x].Value.ToString().Replace(txtSearch.Text, txtReplace.Text);
                }


            }
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            for (int x = Form1.OrganizedSheet.SelectedCells.Count-1; x > 0 ; x--)
            {
                if (Form1.OrganizedSheet.SelectedCells[x].Value != null)
                {
                    if (Form1.OrganizedSheet.SelectedCells[x].Value.ToString().Contains(txtSearch.Text))
                    {
                        
                        Form1.OrganizedSheet.SelectedCells[x].Value = Form1.OrganizedSheet.SelectedCells[x].Value.ToString().Replace(txtSearch.Text, txtReplace.Text);
                        break;
                    }
                }

            }

        }
    }
}
