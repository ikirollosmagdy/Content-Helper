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
 public   class Baby_Clothes
    {
        Database db = new Database();
        TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
        Common_Use common = new Common_Use();
        int Brand = 1, Model = 2, Style = 3, Gender = 4, Colors = 5, Size = 6, Material = 7, 
           Link = 8, Price = 9, Quantity = 10, UnTranslatedCount = 0, ItemConnection = 1;


        private void setupTable()
        {

            DataGridViewComboBoxColumn StyleColumn = new DataGridViewComboBoxColumn();
            StyleColumn.HeaderText = "Category";
            StyleColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            StyleColumn.FlatStyle = FlatStyle.Popup;
            StyleColumn.Items.AddRange("Athletic Wear", "Bodysuit", "Coat", "Costume", "Dress", "Jacket", "Jumpsuit", "Onesie", "Overall", "Pants", "Romper", "Sandal", "Shirt", "Shoe", "Short", "Skirt", "Sleepwear", "Slipper", "Suit", "Swimwear", "Top", "T-Shirt", "Two Pieces Wear", "Underwear");

            DataGridViewComboBoxColumn SizeColumn = new DataGridViewComboBoxColumn();
            SizeColumn.HeaderText = "Size";
            SizeColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            SizeColumn.FlatStyle = FlatStyle.Popup;
            SizeColumn.Items.AddRange("0 - 6 Months", "0 US", "1 UK", "1 US", "1.5 UK", "1.5 US", "1/2 US", "12 - 18 Months", "12 - 24 Months", "12 - 36 Months", "12 Months", "16 EU", "17 EU", "18 - 24 Months", "18 EU", "18.5 EU", "18/19 EU", "19 EU", "19.5 EU", "2 UK", "2 US", "2.5 UK", "2.5 US", "20 EU", "20.5 EU", "20/21 EU", "21 EU", "21/22 EU", "22 EU", "22/23 EU", "23 EU", "23/24 EU", "24 - 30 Months", "24 - 36 Months", "24 EU", "24/25 EU", "25 EU", "25/26 EU", "26 EU", "3 - 6 Months", "3 UK", "3 US", "3.5 UK", "3.5 US", "30 - 36 Months", "4 EU", "4 UK", "4 US", "4.5 EU", "4.5 US", "4/5 US", "5 US", "5.5 US", "6 - 12 Months", "6 - 9 Months", "6 US", "6.5 US", "7 US", "7/8 US", "9 - 12 Months", "9/10 US", "L", "M", "S", "XL");

         


            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("NO", "No.")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Model", "Model")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(StyleColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Gender", "Gender")));

            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Color", "Color")));
        
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(SizeColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("material", "Material")));
          
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
                        for (int y = 1; y < Form1.Sheet.ColumnCount; y++)
                        {
                            if (Form1.Sheet.Columns[y].HeaderText.ToLower().Contains("model") || (Form1.Sheet.Columns[y].HeaderText.ToLower().Contains("code")))
                            {
                                Form1.OrganizedSheet[Model, row].Value = Form1.Sheet[y, row].Value;
                                break;
                            }
                            else
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
                    }
                    catch { }
                 
                    try
                    {

                        Regex regexGender = new Regex(@"([Uu]ni(sex)?)|([Bb]oys?)|([Gg]irls?)");
                        Match matchGender = regexGender.Match(Form1.Sheet.Rows[row].Cells[col].Value.ToString());
                        if (matchGender.Success)
                        {
                           
                                Form1.OrganizedSheet.Rows[row].Cells[Gender].Value = common.getReplacement(matchGender.Value);
                            
                        }


                    }
                    catch (Exception) { }
                    try
                    {
                        if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("color"))
                        {
                            Form1.OrganizedSheet[Colors, row].Value = Form1.Sheet[col, row].Value.ToString().Trim();
                        }
                        else
                        {
                            Regex regexColor = new Regex(@"([Bb]eige|[Bb]lack|[Bb]lue|[Bb]rown|[Cc]lear|[Gg]old|[Gg]reen|[Gg]rey|[Mm]ulti ?[Cc]olor|[Oo]ffWhite|[Oo]range|[Pp]ink|[Pp]urple|[Rr]ed|[Ss]ilver|[Tt]urquoise|[Ww]hite|[Yy]ellow|[Nn]avy|[Ll]ight|[Dd]ark)");
                            MatchCollection matchColor = regexColor.Matches(Form1.Sheet[col, row].Value.ToString());


                            if (matchColor.Count > 1)
                            {
                                Form1.OrganizedSheet[Colors, row].Value = textInfo.ToTitleCase(string.Format("{0} {1}", matchColor[0].Value, matchColor[1].Value)).Trim();

                            }
                            else { Form1.OrganizedSheet[Colors, row].Value = textInfo.ToTitleCase(matchColor[0].Value.Trim()); }
                        }

                    }
                    catch { }
                    //try
                    //{
                    //    if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("size"))
                    //    {
                    //        Form1.OrganizedSheet[Size, row].Value = getSizeBulk(Form1.Sheet[col, row].Value.ToString());
                    //    }
                    //}
                    //catch { }



                    try
                    {
                        if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("material"))
                        {
                            Form1.OrganizedSheet[Material, row].Value = Form1.Sheet[col, row].Value.ToString().Trim();
                        }
                        else
                        {
                            Regex regexMaterial = new Regex(@"([Cc]ot\w+)|([Ll]ea\w+)|([Ww]oo\w+)|([Ss]at\w+)|([Ss]ilk)|([Vv]isco\w+)|([Pp]olye\w+)|([Ll]inen)|([Nn]ylon)|([Ss]uede)");
                            MatchCollection matchMaterial = regexMaterial.Matches(Form1.Sheet[col, row].Value.ToString());


                            if (matchMaterial.Count > 0)
                            {
                                Form1.OrganizedSheet[Material, row].Value = matchMaterial[0].Value.Trim();


                            }
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
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("color", "Color")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Category", "Category")));
           
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Size", "Size")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Gender", "Gender")));


            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArTitle", "Title Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArBran", "Brand Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDes", "Description Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arcolor", "Color Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArtCategory", "Category Arabic")));
         
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArSize", "Size Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArGender", "Gender Arabic")));

            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Link", "Link")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Price", "Price")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Quan", "Quantity")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("itemcon", "Item Connection")));

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
                        setCategory(i);
                        setSize(i);
                        setGender(i);



                        setLink(i);
                        setPrice(i);
                        setQuantity(i);
                        setItemConnection(i);


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
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase(string.Format("{0} {1} {2} {3} for {3} - {4}",
               Form1.OrganizedSheet[Brand, row].Value, Form1.OrganizedSheet[Model, row].Value, Form1.OrganizedSheet[Style, row].Value,
            Form1.OrganizedSheet[Gender, row].Value, Form1.OrganizedSheet[Colors, row].Value));
            Form1.BulkSheet[7, row].Value = string.Format("{0} {1} لل{2} من  {3} - {4}", db.getRecord(Form1.OrganizedSheet[Style, row].Value.ToString()),
                db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()),
                Form1.OrganizedSheet[Model, row].Value, db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString()));
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
            Form1.BulkSheet[2, row].Value = string.Format("<p><b>Product Features:</b></p><ul><li>Brand:{0}</li><li>Model Number:{1}</li><li>Color:{2}</li><li>Size:{3}</li><li>Material:{4}</li><li>Targeted Group:{5}</li></ul>",
                Form1.OrganizedSheet[Brand, row].Value, Form1.OrganizedSheet[Model, row].Value, Form1.OrganizedSheet[Colors, row].Value, Form1.OrganizedSheet[Size, row].Value, Form1.OrganizedSheet[Material, row].Value, Form1.OrganizedSheet[Gender, row].Value);
            Form1.BulkSheet[9, row].Value = string.Format("<p><b>خصائص المنتج:</b></p><ul><li>العلامة  التجارية:{0}</li><li>رقم الموديلr:{1}</li><li>اللون:{2}</li><li>المقاس:{3}</li><li>الخامة:{4}</li><li>المجموعة المستهدفة:{5}</li></ul>",
               db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()), Form1.OrganizedSheet[Model, row].Value, db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString()), Form1.OrganizedSheet[Size, row].Value, db.getRecord(Form1.OrganizedSheet[Material, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString()));
            if (common.CheckEnglish(Form1.BulkSheet[9, row].Value.ToString()))
            {
                Form1.BulkSheet[9, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }


        private void setColor(int row)
        {
            Form1.BulkSheet[3, row].Value = textInfo.ToTitleCase(Form1.OrganizedSheet[Colors, row].Value.ToString().Trim());
            Form1.BulkSheet[10, row].Value = common.getColorMulti(Form1.OrganizedSheet[Colors, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[10, row].Value.ToString()))
            {
                Form1.BulkSheet[10, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }


        private void setCategory(int row)
        {
            Form1.BulkSheet[4, row].Value = getCatBulk(Form1.OrganizedSheet[Style, row].Value.ToString().Trim());
            Form1.BulkSheet[11, row].Value = db.getRecord(Form1.BulkSheet[4, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[11, row].Value.ToString()))
            {
                Form1.BulkSheet[11, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }

        private void setSize(int row)
        {
            Form1.BulkSheet[5, row].Value = Form1.OrganizedSheet[Size, row].Value;
            Form1.BulkSheet[12, row].Value = db.getRecord(Form1.OrganizedSheet[Size, row].Value.ToString());

        }
        private void setGender(int row)
        {
            Form1.BulkSheet[6, row].Value = textInfo.ToTitleCase(Form1.OrganizedSheet[Gender, row].Value.ToString().Trim());
            Form1.BulkSheet[13, row].Value = db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[13, row].Value.ToString()))
            {
                Form1.BulkSheet[13, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
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

        private void setItemConnection(int row)
        {
            try
            {
                if (!Form1.OrganizedSheet[Model, row].Value.ToString().Equals(Form1.OrganizedSheet[Model, row - 1].Value.ToString()))
                {
                    ItemConnection++;
                }
            }
            catch { }
            Form1.BulkSheet[17, row].Value = ItemConnection.ToString();
        }
        private string getCatBulk(string text)
        {
            string value = "";
            if (text.Equals("Sleepwear") || text.Equals("Swimwear") || text.Equals("Underwear")||text.Equals("Two Pieces Wear")
                || text.Equals("Pants")) {
                return text;
            }

            switch (text)
            {
                case "Bodysuit":
                    value = "Bodysuits & Onesies";
                    break;
                case "Onesie":
                    value = "Bodysuits & Onesies";
                    break;
                case "Jacket":
                    value = "Jackets & Coats";
                    break;
                case "Coat":
                    value = "Jackets & Coats";
                    break;
                case "Top":
                    value = "Tops & Shirts";
                    break;
                case "Shirt":
                    value = "Tops & Shirts";
                    break;
                case "T-Shirt":
                    value = "Tops & Shirts";
                    break;


                default:
                    value = text + "s";
                    break;
            }


            return textInfo.ToTitleCase(value.Trim());

        }
    }
}
