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
    public class Snacks_Bakery
    {
        Database db = new Database();
        TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
        Common_Use common = new Common_Use();
        int Brand = 1, Model = 2, Style = 3, Size = 4, Package = 5, Count = 6,
           Link = 7, Price = 8, Quantity = 9, UnTranslatedCount = 0;

        private void setupTable()
        {

            DataGridViewComboBoxColumn StyleColumn = new DataGridViewComboBoxColumn();
            StyleColumn.HeaderText = "Type";
            StyleColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            StyleColumn.FlatStyle = FlatStyle.Popup;
            StyleColumn.Items.AddRange("Biscuits", "Bread", "Crackers", "Cake", "Canned Food", "Chips", "Chocolate", "Dates", "Dried Fruit", "Fruit", "Gum", "Halva", "Jams", "Jelly", "Molokhia", "Nuts", "Corn Oil", "Mixed Oil", "Sunflower Oil", "Oil", "Packaged Meals", "Popcorn", "Soup", "Sweet", "Syrup", "Tahini", "Cookies", "Crackers", "Vegetables", "Spreads", "Honey", "Seeds", "Salt Bakery Products", "Crisps", "Candy");

            DataGridViewComboBoxColumn SizeColumn = new DataGridViewComboBoxColumn();
            SizeColumn.HeaderText = "Total Size";
            SizeColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            SizeColumn.FlatStyle = FlatStyle.Popup;
            SizeColumn.Items.AddRange("1 kg", "10 kg", "100 gm", "1070 gm", "110 gm", "1110 gm", "12 Pieces", "120 gm", "13 gm", "130 gm", "136 gm", "140 gm", "145 gm", "150 gm", "154 gm", "16 oz", "160 gm", "160 ml", "175 gm", "18 Litre", "180 gm", "183 gm", "185 gm", "1850 gm", "19.5 gm", "190 gm", "2 kg", "2.7 Liter", "20 gm", "200 gm", "200 Pieces", "230 gm", "24 Pieces", "240 gm", "25 gm", "25 Pieces", "250 gm", "260 gm", "270 gm", "28 Pieces", "280 gm", "285 gm", "30 gm", "300 gm", "325 gm", "34 gm", "34 Pieces", "35 gm", "350 gm", "36 gm", "360 gm", "380 gm", "40 gm", "400 gm", "40 Pieces", "42 gm", "420 gm", "45 gm", "450 gm", "462 gm", "480 gm", "495 gm", "5 kg", "50 gm", "500 gm", "500 ml", "52 gm", "524 gm", "552 gm", "59 gm", "6 pieces", "60 gm", "600 gm", "610 gm", "650 gm", "660 gm", "70 gm", "700 gm", "710 gm", "72 Pieces", "74 gm", "75 gm", "750 gm", "80 gm", "800 gm", "85 gm", "88 Pieces", "90 gm", "950 gm", "96 gm", "960 gm", "99 gm");

            DataGridViewComboBoxColumn PackingColumn = new DataGridViewComboBoxColumn();
            PackingColumn.HeaderText = "Total Size";
            PackingColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            PackingColumn.FlatStyle = FlatStyle.Popup;
            PackingColumn.Items.AddRange("Basket", "Bottle", "Box", "Can", "Jar", "Pouch");

            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("NO", "No.")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(StyleColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(SizeColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(PackingColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Count", "Count")));

            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Link", "Image Url")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Price", "Price")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Qunatity", "Quantity")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.RowCount = Form1.Sheet.RowCount));
            foreach (DataGridViewColumn co in Form1.OrganizedSheet.Columns)
            {
                co.SortMode = DataGridViewColumnSortMode.NotSortable;
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
                    if (Form1.Sheet.Columns[col].HeaderText.ToLower() == "brand")
                    {
                        Form1.OrganizedSheet[Brand, row].Value = Form1.Sheet[col, row].Value;
                    }

                    try
                    {

                        Regex regexGender = new Regex(@"(Biscuits|Bread|Cake|Canned Food|Chips|Chocolate|Dates|Dried Fruit|Fruit|Gum|Halva|Jams|Jelly|Molokhia|Nuts|Oil|Packaged Meals|Popcorn|Soup|Sweet|Syrup|Tahini|Cookies|Crackers|Vegetables|Spreads|Honey|Seeds|Salt Bakery Products|Crisps|Candy)", RegexOptions.IgnoreCase);
                        Match matchGender = regexGender.Match(Form1.Sheet.Rows[row].Cells[col].Value.ToString());
                        if (matchGender.Success)
                        {

                            Form1.OrganizedSheet.Rows[row].Cells[Style].Value = common.getReplacement(matchGender.Value);

                        }


                    }
                    catch { }



                    if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("price"))
                    {
                        Form1.OrganizedSheet[Price, row].Value = Form1.Sheet[col, row].Value;
                    }
                    if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("link"))
                    {
                        Form1.OrganizedSheet[Link, row].Value = Form1.Sheet[col, row].Value;
                    }
                    if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("qty") || Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("qua"))
                    {
                        Form1.OrganizedSheet[Quantity, row].Value = Form1.Sheet[col, row].Value;
                    }
                }
            }
        }
        public void createBulk()
        {
            UnTranslatedCount = 0;
            Form1.txtUntranslated.GetCurrentParent().Invoke(new Action(() => Form1.txtUntranslated.Text = UnTranslatedCount.ToString()));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Clear()));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Title", "Title")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Bran", "Brand")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Des", "Description")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Type", "Type")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Size", "Total Size")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("package", "Packaging")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("count", "Count")));


            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArTitle", "Title Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArBran", "Brand Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDes", "Description Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArType", "Type Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArSize", "Total Size Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arpackage", "Packaging Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arcount", "Count Arabic")));

            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Link", "Link")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Price", "Price")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Quan", "Quantity")));


            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.RowCount = Form1.OrganizedSheet.RowCount));

            for (int i = 0; i < Form1.OrganizedSheet.RowCount - 1; i++)
            {
                try
                {
                    if (Form1.OrganizedSheet.Rows[i].DefaultCellStyle.BackColor != System.Drawing.Color.HotPink)
                    {

                        setTitle(i);
                        //setBrand(i);
                        //setDescription(i);
                        //setColor(i);
                        //setSize(i);
                        //setCategory(i);

                        //setGender(i);



                        //setLink(i);
                        //setPrice(i);
                        //setQuantity(i);
                        //setItemConnection(i);


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
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase(string.Format("{0} {1}, {2}",
               Form1.OrganizedSheet[Brand, row].Value, Form1.OrganizedSheet[Style, row].Value,
            Form1.OrganizedSheet[Size, row].Value));
            Form1.BulkSheet[7, row].Value = string.Format("{0} من {1}, {2}", db.getRecord(Form1.OrganizedSheet[Style, row].Value.ToString()),
                db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Size, row].Value.ToString()));
            if (common.CheckEnglish(Form1.BulkSheet[7, row].Value.ToString()))
            {
                Form1.BulkSheet[7, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setBrand(int row)
        {
            Form1.BulkSheet[1, row].Value = Form1.OrganizedSheet[Brand, row].Value;
            Form1.BulkSheet[8, row].Value = db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[8, row].Value.ToString()))
            {
                Form1.BulkSheet[8, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setDescription(int row)
        {
            Form1.BulkSheet[2, row].Value = string.Format("<p><b>Product Features:</b></p><ul><li>Brand:{0}</li><li>Type:{1}</li><li>Size:{2}</li><li>Packaging:{3}</li></ul>",
                Form1.OrganizedSheet[Brand, row].Value, Form1.OrganizedSheet[Style, row].Value, Form1.OrganizedSheet[Size, row].Value, Form1.OrganizedSheet[Package, row].Value);
            Form1.BulkSheet[9, row].Value = string.Format("<p><b>خصائص المنتج:</b></p><ul><li>العلامة التجارية:{0}</li><li>النوع:{1}</li><li>الوزن:{2}</li><li>التعبئة:{3}</li></ul>",
               db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Style, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Size, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Package, row].Value.ToString()));
            if (common.CheckEnglish(Form1.BulkSheet[9, row].Value.ToString()))
            {
                Form1.BulkSheet[9, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setType(int row)
        {
            Form1.BulkSheet[3, row].Value = getCatBulk(Form1.OrganizedSheet[Style, row].Value.ToString().Trim());
            Form1.BulkSheet[10, row].Value = db.getRecord(Form1.BulkSheet[3, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[10, row].Value.ToString()))
            {
                Form1.BulkSheet[10, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setSize(int row)
        {
            Form1.BulkSheet[4, row].Value = Form1.OrganizedSheet[Size, row].Value;
            Form1.BulkSheet[11, row].Value = db.getRecord(Form1.OrganizedSheet[Size, row].Value.ToString());

        }

        private void setPackaging(int row)
        {
            Form1.BulkSheet[5, row].Value = Form1.OrganizedSheet[Package, row].Value;
            Form1.BulkSheet[12, row].Value = db.getRecord(Form1.OrganizedSheet[Package, row].Value.ToString());
        }
        private void setCount(int row)
        {
            Form1.BulkSheet[6, row].Value = Form1.OrganizedSheet[Count, row].Value;
            Form1.BulkSheet[13, row].Value = Form1.OrganizedSheet[Count, row].Value;
        }
        private void setLink(int row)
        {
            Form1.BulkSheet[14, row].Value = Form1.OrganizedSheet[Link, row].Value;
        }
        private void setPrice(int row)
        {
            Form1.BulkSheet[15, row].Value = Form1.OrganizedSheet[Price, row].Value;
        }
        private void setQuantity(int row)
        {
            Form1.BulkSheet[16, row].Value = Form1.OrganizedSheet[Quantity, row].Value;
        }


        private string getCatBulk(string text)
        {
            string value = "";


            switch (text)
            {
                case "Biscuits":
                case "Cookies":
                case "Crackers":
                    value = "Biscuits, Cookies & Crackers";
                    break;
                case "Chips":
                case "Crisps":
                    value = "Chips & Crisps";
                    break;
                case "Chocolate":
                case "Candy":
                    value = "Chocolate & Candy";
                    break;
                case "Fruit":
                case "Vegetables":
                    value = "Fruit & Vegetables";
                    break;
                case "Jams":
                case "Spreads":
                case "Honey":
                    value = "Jams, Spreads & Honey";
                    break;
                case "Nuts":
                case "Seeds":
                    value = "Nuts & Seeds";
                    break;
                case "Sweet":
                case "Salt Bakery Products":
                    value = "Sweet & Salt Bakery Products";
                    break;

                default:
                    value = text;
                    break;
            }


            return textInfo.ToTitleCase(value.Trim());

        }
    }
}
