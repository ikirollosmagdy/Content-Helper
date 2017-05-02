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
    public partial class DeveloperForm : Form
    {
        public DeveloperForm()
        {
            InitializeComponent();
        }

        private void btnAddTransMember_Click(object sender, EventArgs e)
        {
            if (txtPassword.Text.Equals(txtRePassword.Text)){
                Database db = new Database();
                db.AddTranslationMember(txtUserName.Text, txtPassword.Text);
                MessageBox.Show("Record added");
            }
            else
            {
                MessageBox.Show("Passwords aren't matched");
            }
        }
    }
}
