using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace helper
{
    public class Adapter
    {
        public void SwitchCategory(object index)
        {
           
            switch (Convert.ToInt32(index))
            {

                case 0:
                    Perfume perfume = new Perfume();
                    perfume.getGender();

                    break;
                case 1:
                    Refrigerator Ref = new Refrigerator();
                    Ref.Organize();
                    break;
                case 2:
                    Makeup makeup = new Makeup();
                    makeup.Organize();
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
                   
                        Perfume per = new Perfume();

                        per.createBulk();
                    

                    break;
                case 1:
                    Refrigerator Ref = new Refrigerator();
                    Ref.creatBulk();
                    break;
                case 2:
                    Makeup makeup = new Makeup();
                    makeup.createBulk();
                    break;

                default:
                    MessageBox.Show("Please choose category first");
                    break;
            }

        }



    }


}
