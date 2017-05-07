using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace helper
{
   public class PowerBank
    {
        Database db = new Database();
        int Brand = 1, Model = 2, Capacity = 3, Colors = 4, Ports = 5, Device = 6,
            Link = 7, Price = 8, Quantity = 9, UnTranslatedCount = 0;

        private void setupTable()
        {
            DataGridViewComboBoxColumn DeviceColumn = new DataGridViewComboBoxColumn();
            DeviceColumn.HeaderText = "Compatible with";
            DeviceColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            DeviceColumn.FlatStyle = FlatStyle.Popup;
            DeviceColumn.Items.AddRange("MP3 Players & MP4 Players", "Multi", "Smart Phones", "Smart Watches", "Tablets");
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("NO", "No.")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Model", "Model")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Capacity", "Capacity")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Color", "Color")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Port", "Port")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(DeviceColumn)));
           
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Link", "Link")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Price", "Price")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Qunatity", "Quantity")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.RowCount = Form1.Sheet.RowCount));
           
                foreach (DataGridViewColumn co in Form1.OrganizedSheet.Columns)
                {
                Form1.OrganizedSheet.Invoke(new Action(() => co.SortMode = DataGridViewColumnSortMode.NotSortable));
                }
            
        }
        public void Organize()
        {
            setupTable();
            int row, col;

            for (row = 0; row < Form1.Sheet.RowCount; row++)
            {
                for (col = 0; col < Form1.Sheet.ColumnCount; col++)
                {
                    try
                    {
                        Form1.OrganizedSheet[0, row].Value = Convert.ToString(row + 1);
                    }
                    catch { }
                    try
                    {
                        for (int y = 0; y < Form1.Sheet.ColumnCount; y++)
                        {
                            Regex regexModel = new Regex(@"(?!\s+)((\w+([-|/|])?)?\w+([-|/|])?\d+(\w+|\d+)?([-|/|])?(\w+|\d+)?)");
                            MatchCollection matchModel = regexModel.Matches(Form1.Sheet[y, row].Value.ToString());


                            if (matchModel.Count > 0)
                            {
                                Form1.OrganizedSheet[Model, row].Value = matchModel[0].Value;
                                break;

                            }
                        }
                    }
                    catch { }




                    try
                    {
                        Regex regexColor = new Regex(@"([Bb]eige|[Bb]lack|[Bb]lue|[Bb]rown|[Cc]lear|[Gg]old|[Gg]reen|[Gg]rey|[Mm]ultiColor|[Oo]ffWhite|[Oo]range|[Pp]ink|[Pp]urple|[Rr]ed|[Ss]ilver|[Tt]urquoise|[Ww]hite|[Yy]ellow)");
                        MatchCollection matchColor = regexColor.Matches(Form1.Sheet[col, row].Value.ToString());


                        if (matchColor.Count > 0)
                        {
                            Form1.OrganizedSheet[Colors, row].Value = matchColor[0].Value;
                           

                        }
                    }
                    catch { }


                    try
                    {
                        Regex regexCapacity= new Regex(@"(\d{4,5})");
                        MatchCollection matchCoapacity = regexCapacity.Matches(Form1.Sheet[col, row].Value.ToString());


                        if (matchCoapacity.Count > 0)
                        {
                            Form1.OrganizedSheet[Capacity, row].Value = matchCoapacity[0].Value;


                        }
                    }
                    catch { }



                }
            }
        }
        public void createBulk() {
            UnTranslatedCount = 0;
            Form1.txtUntranslated.GetCurrentParent().Invoke(new Action(() => Form1.txtUntranslated.Text = UnTranslatedCount.ToString()));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Clear()));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Title", "Title")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Bran", "Brand")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Des", "Description")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Color", "Color")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Device", "Device")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("port", "No of ports")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Model", "Model")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("capacity", "Capacity")));

            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArTitle", "Title Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArBran", "Brand Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDes", "Description Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArColor", "Color Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDevice", "Device Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arport", "No of ports Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArModel", "Model Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arcapacity", "Capacity Arabic")));

            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Link", "Link")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Price", "Price")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Quan", "Quantity")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("EAN", "Suggested EAN")));

            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.RowCount = Form1.OrganizedSheet.RowCount));

            for (int i = 0; i < Form1.OrganizedSheet.RowCount - 1; i++)
            {
                try
                {
                    if (Form1.OrganizedSheet.Rows[i].DefaultCellStyle.BackColor != System.Drawing.Color.HotPink)
                    {
                        setTitle(i);
                        setBrand(i);
                      setDescription(i);
                       setColor(i);
                         setDevice(i);
                      setPorts(i);
                       setModel(i);
                        setCapacity(i);
                          setLink(i);
                           setPrice(i);
                         setQuantity(i);
                        Form1.txtUntranslated.GetCurrentParent().Invoke(new Action(() => Form1.txtUntranslated.Text = UnTranslatedCount.ToString()));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

        }
        private void setTitle(int row)
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            string title = "", ArTitle = "";
            if (Form1.OrganizedSheet[Brand, row].Value.ToString().ToLower() != "other")
            {
                title = Form1.OrganizedSheet[Brand, row].Value + " " + Form1.OrganizedSheet[Model, row].Value + " ";
                ArTitle = " من " + db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()) + " " + Form1.OrganizedSheet[Model, row].Value;
            }
            title = title + "Power bank, " + Form1.OrganizedSheet[Capacity, row].Value + " mAh, " + Form1.OrganizedSheet[Colors, row].Value;
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase(title);
            Form1.BulkSheet[8, row].Value = "بطارية شحن محمول " + ArTitle + "، " + Form1.OrganizedSheet[Capacity, row].Value + " مللي امبير، " +
                db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString());
            if (db.CheckEnglish(Form1.BulkSheet[8, row].Value.ToString()))
            {
                Form1.BulkSheet[8, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setBrand(int row)
        {
            Form1.BulkSheet[1, row].Value = Form1.OrganizedSheet[Brand, row].Value;
            Form1.BulkSheet[9, row].Value = db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString());
            if (db.CheckEnglish(Form1.BulkSheet[9, row].Value.ToString()))
            {
                Form1.BulkSheet[9, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setDescription(int row)
        {
            Form1.BulkSheet[2, row].Value = "<ul> <li>Brand :" + Form1.OrganizedSheet[Brand, row].Value + "</li> <li>Color :" +
                Form1.OrganizedSheet[Colors, row].Value + "</li> <li>Capacity :" + Form1.OrganizedSheet[Capacity, row].Value +
                "mAh</li> <li>Number of Ports :" + Form1.OrganizedSheet[Ports, row].Value + "</li> <li>Compatible with :" +
                Form1.OrganizedSheet[Device, row].Value + "</li> </ul>";

            Form1.BulkSheet[10,row].Value= "<ul> <li>العلامة التجارية :" + db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()) + "</li> <li>اللون :" +
              db.getRecord(  Form1.OrganizedSheet[Colors, row].Value.ToString()) + "</li> <li>السعة :" + db.getRecord(Form1.OrganizedSheet[Capacity, row].Value.ToString()) +
                "مللي امبير</li> <li>عدد المنافذ :" +db.getRecord( Form1.OrganizedSheet[Ports, row].Value.ToString()) + "</li> <li>متوافق مع :" +
               db.getRecord( Form1.OrganizedSheet[Device, row].Value.ToString()) + "</li> </ul>";
            if (db.CheckEnglish(Form1.BulkSheet[10, row].Value.ToString()))
            {
                Form1.BulkSheet[10, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
      private  void setColor(int row)
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            Form1.BulkSheet[3, row].Value = textInfo.ToTitleCase(Form1.OrganizedSheet[Colors, row].Value.ToString());
            Form1.BulkSheet[11,row].Value=db.getRecord(Form1.OrganizedSheet[Colors,row].Value.ToString());
            if (db.CheckEnglish(Form1.BulkSheet[11, row].Value.ToString()))
            {
                Form1.BulkSheet[11, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setDevice(int row)
        {
            Form1.BulkSheet[4, row].Value = Form1.OrganizedSheet[Device, row].Value;
            Form1.BulkSheet[12, row].Value = db.getRecord(Form1.OrganizedSheet[Device, row].Value.ToString());
            if (db.CheckEnglish(Form1.BulkSheet[12, row].Value.ToString()))
            {
                Form1.BulkSheet[12, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setPorts(int row)
        {
            Form1.BulkSheet[5, row].Value = Form1.OrganizedSheet[Ports, row].Value;
            Form1.BulkSheet[13,row].Value= Form1.OrganizedSheet[Ports, row].Value;
        }
        private void setModel(int row)
        {
            Form1.BulkSheet[6, row].Value = Form1.OrganizedSheet[Model, row].Value;
            Form1.BulkSheet[14, row].Value = Form1.OrganizedSheet[Model, row].Value;
        }
        private void setCapacity(int row)
        {
            Form1.BulkSheet[7, row].Value = getCapacityAttribute(Form1.OrganizedSheet[Capacity, row]);
            Form1.BulkSheet[15, row].Value = db.getRecord(Form1.BulkSheet[7, row].Value.ToString());
            if (db.CheckEnglish(Form1.BulkSheet[15, row].Value.ToString()))
            {
                Form1.BulkSheet[15, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setLink(int row)
        {
            Form1.BulkSheet[16, row].Value = Form1.OrganizedSheet[Link, row].Value;
        }

        private void setPrice(int row)
        {
            Form1.BulkSheet[17, row].Value = Form1.OrganizedSheet[Price, row].Value;
        }
        private void setQuantity(int row)
        {
            Form1.BulkSheet[18, row].Value = Form1.OrganizedSheet[Quantity, row].Value;
        }

      
        private string getCapacityAttribute(DataGridViewCell cell)
        {
            string value = "";
            int cap = Convert.ToInt32(cell.Value.ToString());
            if (cap > 3000 && cap <= 4000)
            {
                value = "3001-4000mAh";
            } else if (cap > 4000 && cap <= 5000)
            {
                value = "4001-5000mAh";
            }
            else if (cap > 5000 && cap <= 6000)
            {
                value = "5001-6000mAh";
            }
            else if (cap > 6000 && cap <= 7000)
            {
                value = "6001-7000mAh";
            }
            else if (cap > 7000 && cap <= 8000)
            {
                value = "7001-8000mAh";

            }
            else if (cap > 8000 && cap <= 9000)
            {
                value = "9001-10000mAh";
            }
            else if (cap > 10000 && cap <= 15000)
            {
                value = "10001-15000mAh";
            }else if (cap > 15000 && cap <= 20000)
            {
                value = "15001-20000mAh";
            
            }else if(cap <= 3000)
            {
                value = "Less Than 3001mAh";
            }else if(cap >= 20000)
            {
                value = "20000mAh & Above";
            }

            return value;
        }

    }
}
