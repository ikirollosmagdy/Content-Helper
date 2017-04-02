using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace test
{
    class Adapter
    {
        public void SwitchCategory(object index)
        {

            switch (Convert.ToInt32(index))
            {

                case 0:
                    Perfume.getGender();
                    ;
                    
                    break;
                default:
                    MessageBox.Show("Please choose category first");
                    break;
            }
            Form1.txtStats.GetCurrentParent().Invoke(new Action(() => Form1.txtStats.Text = "Finished"));
            Form1.PBar.GetCurrentParent().Invoke(new Action(() => Form1.PBar.Visible = false));
            
        }

        public void SwitchBulk(object index)
        {
            switch (Convert.ToInt32(index))
            {

                case 0:
                    Perfume.createBulk();
                    break;
                default:
                    MessageBox.Show("Please choose category first");
                    break;
            }
        }


    }
}
