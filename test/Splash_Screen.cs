using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace helper
{
    public partial class Splash_Screen : Form
    {
        public Splash_Screen()
        {
            InitializeComponent();
        }

        private void btnChooseLang_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.IsEgypt = RadEgypt.Checked;
            Properties.Settings.Default.Save();
            Close();
        }

        private void Splash_Screen_Activated(object sender, EventArgs e)
        {
           if(Properties.Settings.Default.IsEgypt)
            {
                RadEgypt.Checked = true;
            }
            else
            {
                RadUAE.Checked = true;
            }
        }
    }
}
