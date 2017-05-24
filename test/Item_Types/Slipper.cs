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
  public  class Slipper
    {
        Database db = new Database();
        TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
        Common_Use common = new Common_Use();
        int Brand = 1, Model = 2, Type = 3, Gender = 4, Colors = 5, Size = 6, Material = 7,
           Link = 8, Price = 8, Quantity = 10, UnTranslatedCount = 0, ItemConnection = 1;
        private void setupTable()
        {

            DataGridViewComboBoxColumn TypeColumn = new DataGridViewComboBoxColumn();
            TypeColumn.HeaderText = "Slipper Type";
            TypeColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            TypeColumn.FlatStyle = FlatStyle.Popup;
            TypeColumn.Items.AddRange("Clog", "Comfort & Medical", "Flip Flops", "Moccasin", "Mule", "Slides", "Thong", "Warm Slippers", "Wedges");

            DataGridViewComboBoxColumn SizeColumn = new DataGridViewComboBoxColumn();
            SizeColumn.HeaderText = "Size";
            SizeColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            SizeColumn.FlatStyle = FlatStyle.Popup;
            SizeColumn.Items.AddRange("10 UK", "10 US", "10 US", "10.5 UK", "10.5 US", "11 UK", "11 UK", "11 US", "11.5 UK", "11.5 US", "11.5 US", "12 UK", "12 US", "12 US", "12.5 UK", "12.5 US", "13 UK", "13 US", "13.5 UK", "13.5 US", "14 UK", "14 US", "14.5 UK", "14.5 US", "15 UK", "15 US", "27 EU", "28 EU", "29 EU", "3 UK", "3 US", "3.5 UK", "3.5 US", "30 EU", "31 EU", "32 EU", "33 EU", "34 EU", "35 EU", "35.5 EU", "36 2/3 EU", "36 EU", "36.5 EU", "36.7 EU", "37 1/3 EU", "37 EU", "37.5 EU", "38 2/3 EU", "38 EU", "38.5 EU", "39 1/3 EU", "39 EU", "39.5 EU", "4 UK", "4 US", "4.5 UK", "4.5 US", "40 2/3 EU", "40 EU", "40.5 EU", "41 1/3 EU", "41 EU", "41.5 EU", "42 2/3 EU", "42 EU", "42.5 EU", "43 1/3 EU", "43 EU", "43.3 EU", "43.5 EU", "44 2/3 EU", "44 EU", "44.5 EU", "45 EU", "45.5 EU", "46 EU", "46.5 EU", "47 EU", "48 EU", "48.5 EU", "49 EU", "5 UK", "5 US", "5.5 UK", "5.5 US", "6 UK", "6 US", "6.5 UK", "6.5 US", "7 UK", "7 US", "7.5 UK", "7.5 US", "7.7 US", "8 UK", "8 US", "8.5 UK", "8.5 US", "9 UK", "9 US", "9.5 UK", "9.5 US");

           


            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("NO", "No.")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Model", "Model")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(TypeColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Gender", "Gender")));

            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Color", "Color")));
            //  Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("size", "Size")));
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
                        Regex regexType = new Regex(@"([Cc]log|[Ff]lip ?[Ff]lops?|[Ss]lides?|[Ss]lippers?|[Mm]occasin)");
                        Match matchType = regexType.Match(Form1.Sheet[col, row].Value.ToString());
                        if (matchType.Success)
                        {
                            Form1.OrganizedSheet[Type, row].Value = common.getReplacement(matchType.Value);

                        }
                    }
                    catch { }
                    try
                    {

                        Regex regexGender = new Regex(@"(([Ww]o)?[Mm][ae]n)|([Uu]ni(sex)?)|([Bb]oys?)|([Gg]irls?)");
                        Match matchGender = regexGender.Match(Form1.Sheet.Rows[row].Cells[col].Value.ToString());
                        if (matchGender.Success)
                        {
                            if (matchGender.Value.ToLower() == "woman")
                            {
                                Form1.OrganizedSheet.Rows[row].Cells[Gender].Value = "Women";
                            }
                            else
                            {
                                Form1.OrganizedSheet.Rows[row].Cells[Gender].Value = common.getReplacement(matchGender.Value);
                            }
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
                            Regex regexColor = new Regex(@"([Bb]eige|[Bb]lack|[Bb]lue|[Bb]rown|[Cc]lear|[Gg]old|[Gg]reen|[Gg]rey|[Mm]ulti ?[Cc]olor|[Oo]ffWhite|[Oo]range|[Pp]ink|[Pp]urple|[Rr]ed|[Ss]ilver|[Tt]urquoise|[Ww]hite|[Yy]ellow)");
                            MatchCollection matchColor = regexColor.Matches(Form1.Sheet[col, row].Value.ToString());


                            if (matchColor.Count > 1)
                            {
                                Form1.OrganizedSheet[Colors, row].Value = textInfo.ToTitleCase(string.Format("{0} {1}", matchColor[0].Value, matchColor[1].Value)).Trim();

                            }
                            else { Form1.OrganizedSheet[Colors, row].Value = textInfo.ToTitleCase(matchColor[0].Value.Trim()); }
                        }

                    }
                    catch { }
                    try
                    {
                        if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("size"))
                        {
                            Form1.OrganizedSheet[Size, row].Value = getSizeBulk(Form1.Sheet[col, row].Value.ToString());
                        }
                    }
                    catch { }



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
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("size", "Size")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Slipper", "Slipper Type")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("gender", "Gender")));
            

            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArTitle", "Title Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArBran", "Brand Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDes", "Description Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arcolor", "Color Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arsize", "Size Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArStyle", "Slipper Type Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Argender", "Gender Arabic")));
            

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
                        setType(i);
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
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase(string.Format("{0} {1} {2} for {4} - {5}",
               Form1.OrganizedSheet[Brand, row].Value, Form1.OrganizedSheet[Model, row].Value, Form1.OrganizedSheet[Type, row].Value,
             Form1.OrganizedSheet[Gender, row].Value, Form1.OrganizedSheet[Colors, row].Value));
            Form1.BulkSheet[7, row].Value = string.Format("{0} {1} لل{2} من  {4} - {5}", db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString()),
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
            Form1.BulkSheet[10, row].Value = db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[10, row].Value.ToString()))
            {
                Form1.BulkSheet[10, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }


        private void setSize(int row)
        {
            Form1.BulkSheet[4, row].Value = Form1.OrganizedSheet[Size, row].Value;
            Form1.BulkSheet[11, row].Value = Form1.OrganizedSheet[Size, row].Value;

        }
        private void setType(int row)
        {
            Form1.BulkSheet[5, row].Value = textInfo.ToTitleCase(Form1.OrganizedSheet[Type, row].Value.ToString().Trim());
            Form1.BulkSheet[12, row].Value = db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[12, row].Value.ToString()))
            {
                Form1.BulkSheet[12, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
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
        private string getSizeBulk(string text)
        {
            string value = "";

            try
            {
                int si = Convert.ToInt32(text);
                if (si > 0)
                {
                    if (si > 27 && si < 57)
                    {
                        value = si + " EU";
                    }
                    else
                    {
                        value = "";
                    }
                }
            }
            catch { }
            return value.Trim();

        }

    }
}
