using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Globalization;

namespace helper
{
  public  class Handbags
    {
        Database db = new Database();
        TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
        Common_Use common = new Common_Use();
        int Brand = 1, Model = 2, Type = 3, Gender = 4, Colors = 5,  Material = 6, 
           Link = 7, Price = 8, Quantity = 9, UnTranslatedCount = 0, ItemConnection = 1;


        private void setupTable()
        {

            DataGridViewComboBoxColumn TypeColumn = new DataGridViewComboBoxColumn();
            TypeColumn.HeaderText = "Type";
            TypeColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            TypeColumn.FlatStyle = FlatStyle.Popup;
            TypeColumn.Items.AddRange("Baguette Bag", "Briefcase", "Bucket Bag", "Clutch", "Crossbody Bag", "Duffle Handbag", "Flap Bag", "Frame Bag", "Handbag Accessories", "Handbags Set", "Hobo", "Messenger Bag", "Saddle Bag", "Satchels Bag", "Shopper Bag", "Tote Bag");


            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("NO", "No.")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Model", "Model")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(TypeColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Gender", "Gender")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Color", "Color")));
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
                        Regex regexType = new Regex(@"([Bb]riefcase|[Hh]obo|[Cc]lutch)");
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
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("material", "Material")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("type", "Type")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Gender", "Gender")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("familey", "Color Family")));
            
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArTitle", "Title Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArBran", "Brand Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDes", "Description Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arcolor", "Color Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Armaterial", "Material Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Artype", "Type Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArGender", "Gender Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arfamiley", "Color Family Arabic")));

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
                       setMaterial(i);
                        setType(i);
                        setGender(i);
                        setFamilyColor(i);

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
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase(string.Format("{0} {1} {2} for {3} - {4}",
               Form1.OrganizedSheet[Brand, row].Value, Form1.OrganizedSheet[Model, row].Value, Form1.OrganizedSheet[Type, row].Value,
             Form1.OrganizedSheet[Gender, row].Value, Form1.OrganizedSheet[Colors, row].Value));
            Form1.BulkSheet[8, row].Value = string.Format("{0} لل{1} من {2} {3} - {4}", db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString()),
                db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()),
                Form1.OrganizedSheet[Model, row].Value, db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString()));
            if (common.CheckEnglish(Form1.BulkSheet[8, row].Value.ToString()))
            {
                Form1.BulkSheet[8, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setBrand(int row)
        {
            Form1.BulkSheet[1, row].Value = Form1.OrganizedSheet[Brand, row].Value;
            Form1.BulkSheet[9, row].Value = db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[9, row].Value.ToString()))
            {
                Form1.BulkSheet[9, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setDescription(int row)
        {
            Form1.BulkSheet[2, row].Value = string.Format("<p><b>Product Features:</b></p><ul><li>Brand:{0}</li><li>Model Number:{1}</li><li>Color:{2}</li><li>Type:{3}</li><li>Material:{4}</li><li>Targeted Group:{5}</li></ul>",
                Form1.OrganizedSheet[Brand, row].Value, Form1.OrganizedSheet[Model, row].Value, Form1.OrganizedSheet[Colors, row].Value, Form1.OrganizedSheet[Type, row].Value, Form1.OrganizedSheet[Material, row].Value, Form1.OrganizedSheet[Gender, row].Value);
            Form1.BulkSheet[10, row].Value = string.Format("<p><b>خصائص المنتج:</b></p><ul><li>العلامة  التجارية:{0}</li><li>رقم الموديلr:{1}</li><li>اللون:{2}</li><li>نوع المنتج:{3}</li><li>الخامة:{4}</li><li>المجموعة المستهدفة:{5}</li></ul>",
               db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()), Form1.OrganizedSheet[Model, row].Value, db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString()), Form1.OrganizedSheet[Type, row].Value, db.getRecord(Form1.OrganizedSheet[Material, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString()));
            if (common.CheckEnglish(Form1.BulkSheet[10, row].Value.ToString()))
            {
                Form1.BulkSheet[10, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }


        private void setColor(int row)
        {
            Form1.BulkSheet[3, row].Value = textInfo.ToTitleCase(Form1.OrganizedSheet[Colors, row].Value.ToString().Trim());
            Form1.BulkSheet[11, row].Value = db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[11, row].Value.ToString()))
            {
                Form1.BulkSheet[11, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }

        private void setMaterial(int row)
        {
            Form1.BulkSheet[4, row].Value = getMaterial(Form1.OrganizedSheet[Material, row].Value.ToString());
            Form1.BulkSheet[12, row].Value = db.getRecord(Form1.BulkSheet[4, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[12, row].Value.ToString()))
            {
                Form1.BulkSheet[12, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }

        private void setType(int row)
        {
            Form1.BulkSheet[5, row].Value = getTypeAttribute(Form1.OrganizedSheet[Type,row].Value.ToString());
            Form1.BulkSheet[13, row].Value = db.getRecord(Form1.BulkSheet[5, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[13, row].Value.ToString()))
            {
                Form1.BulkSheet[13, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }

        private void setGender(int row)
        {
            Form1.BulkSheet[6, row].Value = textInfo.ToTitleCase(Form1.OrganizedSheet[Gender, row].Value.ToString().Trim());
            Form1.BulkSheet[14, row].Value = db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[14, row].Value.ToString()))
            {
                Form1.BulkSheet[14, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }

        private void setFamilyColor(int row)
        {

            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            string value = Form1.OrganizedSheet[Colors, row].Value.ToString();
            int WordCount = value.Split(' ').Length;
            if (WordCount > 1)
            {
                if (value.ToLower().Contains("light") || value.ToLower().Contains("dark"))
                {
                    Form1.BulkSheet[7, row].Value = common.getReplacement(value);
                }
                else
                {
                    Form1.BulkSheet[7, row].Value = "Multi Color";
                }
            }
            else
            {
                Form1.BulkSheet[7, row].Value = common.getReplacement(Form1.OrganizedSheet[Colors, row].Value.ToString());
            }

            Form1.BulkSheet[15, row].Value = db.getRecord(Form1.BulkSheet[6, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[14, row].Value.ToString()))
            {
                Form1.BulkSheet[15, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[15, row].Style.BackColor = Color.Empty;
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
            Form1.BulkSheet[19, row].Value = ItemConnection.ToString();
        }
        private string getTypeAttribute(string text)
        {
            string value = "";
            switch (text)
            {
                case "Handbag Accessories":
                    value = text;
                    break;
                case "Clutch":
                    value = text + "es";
                    break;

                default:
                    value = text+"s";
                    break;
            }
            return value;
        }
        private string getMaterial(string text)
        {
            string value = "";
            Regex regmat = new Regex(@"([Cc]ot\w+)|([Ll]ea\w+)|([Ww]oo\w+)|([Ss]at\w+)|([Ss]ilk)|([Vv]isco\w+)|([Pp]olye\w+)|([Ll]inen)|([Nn]ylon)|([Vv]elvet)|(pu)",RegexOptions.IgnoreCase);
            MatchCollection matchmat = regmat.Matches(text);
            if (matchmat.Count > 1)
            {
                value = "Mixed";
            }
            else
            {
                value = textInfo.ToTitleCase(matchmat[0].Value);
            }
            return value.Trim();
        }
    }
}
