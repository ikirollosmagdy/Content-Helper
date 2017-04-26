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
                    case 3:
                    Mobile_Accessories acc = new Mobile_Accessories();
                    acc.Organize();
                    break;

                case 4:
                    Cables cable = new Cables();
                    cable.Organize();
                    break;


                default:
                    MessageBox.Show("Please choose category first");
                    break;
            }
            Form1.txtStats.GetCurrentParent().Invoke(new Action(() => Form1.txtStats.Text = "Finished"));
            Form1.PBar.GetCurrentParent().Invoke(new Action(() => Form1.PBar.Style = ProgressBarStyle.Continuous));

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
                case 3:
                    Mobile_Accessories acc = new Mobile_Accessories();
                    acc.createBulk();
                    break;
                case 4:
                    Cables cable = new Cables();
                    cable.createBulk();
                    break;

                default:
                    MessageBox.Show("Please choose category first");
                    break;
            }

        }



    }


}
