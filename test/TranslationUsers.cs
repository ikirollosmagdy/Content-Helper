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
    public partial class TranslationUsers : Form
    {
        public TranslationUsers()
        {
            InitializeComponent();
        }

        private void btnLoginTransUser_Click(object sender, EventArgs e)
        {
            Database db = new Database();

           
            
            if (txtTransPassword.Text.Equals(db.getUserPassword(txtTransUserName.Text)))
            {
                if (!Form1.tabs.TabPages.Contains(Form1.translationTab))
                {
                    Form1.tabs.TabPages.Add(Form1.translationTab);
                    Close();
                }
                else
                {
                    Close();
                }
            }
            else
            {
                MessageBox.Show("Username or Password is incorrect");
            }
            
        }

        private void TranslationUsers_Load(object sender, EventArgs e)
        {
            txtTransUserName.Focus();
            txtTransUserName.Select();
        }
    }
}
