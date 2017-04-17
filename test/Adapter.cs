using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace helper
{
    class Adapter
    {
        public void SwitchCategory(object index)
        {
            clearSheets();
            switch (Convert.ToInt32(index))
            {

                case 0:
                    Perfume.getGender();
                                      
                    break;
                case 1:
                    Refrigerator Ref= new Refrigerator();
                    Ref.Organize();
                    break;

                default:
                    MessageBox.Show("Please choose category first");
                    break;
            }
            Form1.txtStats.GetCurrentParent().Invoke(new Action(() => Form1.txtStats.Text = "Finished"));
            Form1.PBar.GetCurrentParent().Invoke(new Action(() => Form1.PBar.Visible = false));
            
        }
        public void clearSheets()
        {
            
                int row, col;

            
                for (row = 0; row < Form1.BulkSheet.RowCount; row++)
                {
                    for (col = 0; col < Form1.BulkSheet.ColumnCount; col++)
                    {
                        Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet[col, row].Value = ""));
                    }
                }
               
           
        }

        public void SwitchBulk(object index)
        {
          
            switch (Convert.ToInt32(index))
            {

                case 0:
                    Perfume.createBulk();
                    break;
                case 1:
                    Refrigerator Ref = new Refrigerator();
                    Ref.creatBulk();
                    break;

                default:
                    MessageBox.Show("Please choose category first");
                    break;
            }
        }
       


    }
}
