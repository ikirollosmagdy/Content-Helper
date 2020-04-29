using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FuzzyString;

namespace helper
{
    public partial class New_item_form : Form
    {
        public New_item_form()
        {
            InitializeComponent();
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Space)
            {
                label1.Text = ((1-ComparisonMetrics.JaccardDistance(textBox1.Text.Trim(), "Basem Abduallah"))*100).ToString();
            }
        }
    }
}
