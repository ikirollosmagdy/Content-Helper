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
   public class Tops
    {
        Database db = new Database();
        TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
        Common_Use common = new Common_Use();
        int Brand = 1, Model = 2, Type = 3,  Gender = 4, Colors = 5, Size = 6,Material=7,Sleeve=8,NeckStyle=9,
           Link = 10, Price = 11, Quantity = 12, UnTranslatedCount = 0,ItemConnection=1;
        private void setupTable()
        {
            
            DataGridViewComboBoxColumn TypeColumn = new DataGridViewComboBoxColumn();
            TypeColumn.HeaderText = "Top Style";
            TypeColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            TypeColumn.FlatStyle = FlatStyle.Popup;
            TypeColumn.Items.AddRange("Asymmetrical Top", "Backless Top", "Blouse", "Bodysuit", "Crop Top", "Vest", "Henley Top", "Hoody", "Sweatshirt", "Long Line Top", "Polo", "Pullover", "Shirt", "T-Shirt", "Tank Top", "Tube Top", "Two In One Top", "Two Pieces Wear", "Wrap Top");

            DataGridViewComboBoxColumn SizeColumn = new DataGridViewComboBoxColumn();
            SizeColumn.HeaderText = "Size";
            SizeColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            SizeColumn.FlatStyle = FlatStyle.Popup;
            SizeColumn.Items.AddRange("0 US", "10 - 11 Years", "10 UK", "10 US", "11 - 12 Years", "12 - 13 Years", "12 UK", "12 US", "13 - 14 Years", "14 - 15 Years", "14 - 16 UK", "14 UK", "14 US", "15 - 16 Years", "16 UK", "16 US", "18 - 20 UK", "18 UK", "18 US", "2 US", "20 UK", "20 US", "22 - 24 UK", "22 UK", "22 US", "24 UK", "24 US", "26 - 28 UK", "26 UK", "28 EU", "28 UK", "29 EU", "3 - 4 Years", "30 - 32 UK", "30 EU", "30 UK", "31 EU", "32 EU", "32 UK", "33 EU", "34 EU", "34 UK", "36 EU", "36 UK", "38 EU", "38 UK", "39 EU", "3XL", "4 - 5 Years", "4 UK", "4 US", "40 EU", "40 UK", "41 EU", "42 EU", "42 UK", "43 EU", "44 EU", "44 UK", "45 EU", "46 EU", "46 UK", "48 EU", "48 UK", "4XL", "5 - 6 Years", "5 Years", "50 EU", "50 UK", "52 EU", "54 EU", "56 EU", "5XL", "6 - 7 Years", "6 UK", "6 US", "6XL", "7 - 8 Years", "7XL", "8 - 9 Years", "8 UK", "8 US", "8XL", "9 - 10 Years", "Free Size", "L", "L/XL", "M", "M/L", "S", "S/M", "UK 10", "XL", "XL/XXL", "XS", "XXL", "XXS");

            DataGridViewComboBoxColumn SleeveColumn = new DataGridViewComboBoxColumn();
            SleeveColumn.HeaderText = "Sleeve Long";
            SleeveColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            SleeveColumn.FlatStyle = FlatStyle.Popup;
            SleeveColumn.Items.AddRange("Full Sleeve", "Short Sleeve", "Single Sleeve", "Sleeveless", "Three Quarter Sleeve");

            DataGridViewComboBoxColumn NickStyleColumn = new DataGridViewComboBoxColumn();
            NickStyleColumn.HeaderText = "Neck Style";
            NickStyleColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            NickStyleColumn.FlatStyle = FlatStyle.Popup;
            NickStyleColumn.Items.AddRange("Asymmetrical Neck", "Cowl Neck", "Halter Neck", "High Neck", "Mixed Neck", "Off Shoulder", "Round Neck", "Round Split Neck", "Shirt Neck", "Square Neck", "Sweetheart Neck", "V Neck");


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
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(SleeveColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(NickStyleColumn)));
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
                        Regex regexType = new Regex(@"((([Ss]weat))?([Tt]-?)?[Ss]hirt)|([Pp]olo)|([Bb]louse)|([Pp]ullover)|([Hh]ood\w+)");
                        Match matchType = regexType.Match(Form1.Sheet[col, row].Value.ToString());
                        if (matchType.Success)
                        {
                            Form1.OrganizedSheet[Type, row].Value = common.getReplacement(matchType.Value);
                            if(Form1.OrganizedSheet[Type, row].Value.ToString().Equals("Polo"))
                            {
                                Form1.OrganizedSheet[NeckStyle, row].Value = "Shirt Neck";
                            }
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
                            Regex regexMaterial = new Regex(@"([Cc]ot\w+)|([Ll]ea\w+)|([Ww]oo\w+)|([Ss]at\w+)|([Ss]ilk)|([Vv]isco\w+)|([Pp]olye\w+)|([Ll]inen)|([Nn]ylon)");
                            MatchCollection matchMaterial = regexMaterial.Matches(Form1.Sheet[col, row].Value.ToString());


                            if (matchMaterial.Count > 0)
                            {
                                Form1.OrganizedSheet[Material, row].Value = common.getReplacement(matchMaterial[0].Value.Trim());


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
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("fabType", "Fabric Type")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("topStyle", "Top Style")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("size", "Size")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("gender", "Gender")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Sleeve", "Sleeve length")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("neck", "Neck Style")));

            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArTitle", "Title Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArBran", "Brand Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDes", "Description Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arcolor", "Color Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArfabType", "Fabric Type Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArtopStyle", "Top Style Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arsize", "Size Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Argender", "Gender Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArSleeve", "Sleeve length Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arneck", "Neck Style Arabic")));

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
                     setTopStyle(i);
                       setSize(i);
                    setGender(i);
                        setSleevelength(i);
                        setNeckStyle(i);
           
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
               Form1.OrganizedSheet[Brand,row].Value, Form1.OrganizedSheet[Model, row].Value, Form1.OrganizedSheet[Type, row].Value,
               Form1.OrganizedSheet[Gender, row].Value, Form1.OrganizedSheet[Colors, row].Value));
            Form1.BulkSheet[10, row].Value = string.Format("{0} لل{1} من {2} {3} - {4}",db.getRecord(Form1.OrganizedSheet[Type,row].Value.ToString()),
                db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()),
                Form1.OrganizedSheet[Model,row].Value, db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString()));
            if (common.CheckEnglish(Form1.BulkSheet[10, row].Value.ToString()))
            {
                Form1.BulkSheet[10, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setBrand(int row)
        {
            Form1.BulkSheet[1, row].Value = Form1.OrganizedSheet[Brand, row].Value;
            Form1.BulkSheet[11, row].Value = db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[11, row].Value.ToString()))
            {
                Form1.BulkSheet[11, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setDescription(int row)
        {
            Form1.BulkSheet[2, row].Value = string.Format("<p><b>Product Features:</b></p><ul><li>Brand:{0}</li><li>Model Number:{1}</li><li>Color:{2}</li><li>Size:{3}</li><li>Material:{4}</li><li>Targeted Group:{5}</li></ul>",
                Form1.OrganizedSheet[Brand,row].Value, Form1.OrganizedSheet[Model, row].Value, Form1.OrganizedSheet[Colors, row].Value, Form1.OrganizedSheet[Size, row].Value, Form1.OrganizedSheet[Material, row].Value, Form1.OrganizedSheet[Gender, row].Value);
            Form1.BulkSheet[12, row].Value = string.Format("<p><b>خصائص المنتج:</b></p><ul><li>العلامة  التجارية:{0}</li><li>رقم الموديلr:{1}</li><li>اللون:{2}</li><li>المقاس:{3}</li><li>الخامة:{4}</li><li>المجموعة المستهدفة:{5}</li></ul>",
               db.getRecord(  Form1.OrganizedSheet[Brand, row].Value.ToString()), Form1.OrganizedSheet[Model, row].Value, db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString()), Form1.OrganizedSheet[Size, row].Value, db.getRecord(Form1.OrganizedSheet[Material, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString()));
            if (common.CheckEnglish(Form1.BulkSheet[12, row].Value.ToString()))
            {
                Form1.BulkSheet[12, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setColor (int row)
        {
            Form1.BulkSheet[3, row].Value =textInfo.ToTitleCase( Form1.OrganizedSheet[Colors, row].Value.ToString().Trim());
            Form1.BulkSheet[13,row].Value=db.getRecord( Form1.OrganizedSheet[Colors,row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[13, row].Value.ToString()))
            {
                Form1.BulkSheet[13, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setMaterial(int row)
        {
            Form1.BulkSheet[4, row].Value = getMaterial(Form1.OrganizedSheet[Material, row].Value.ToString());
            Form1.BulkSheet[14, row].Value = db.getRecord(Form1.BulkSheet[4, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[14, row].Value.ToString()))
            {
                Form1.BulkSheet[14, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setTopStyle(int row)
        {
            Form1.BulkSheet[5, row].Value = getTopStyle(Form1.OrganizedSheet[Type, row].Value.ToString());
            Form1.BulkSheet[15, row].Value = db.getRecord(Form1.BulkSheet[5, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[15, row].Value.ToString()))
            {
                Form1.BulkSheet[15, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }

        }
        private void setSize(int row) {
            Form1.BulkSheet[6, row].Value = Form1.OrganizedSheet[Size, row].Value;
            Form1.BulkSheet[16, row].Value = db.getRecord(Form1.OrganizedSheet[Size, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[16, row].Value.ToString()))
            {
                Form1.BulkSheet[16, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setGender(int row) {
            Form1.BulkSheet[7, row].Value =textInfo.ToTitleCase( Form1.OrganizedSheet[Gender, row].Value.ToString().Trim());
            Form1.BulkSheet[17, row].Value = db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[17, row].Value.ToString()))
            {
                Form1.BulkSheet[17, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setSleevelength(int row) {
            Form1.BulkSheet[8, row].Value = Form1.OrganizedSheet[Sleeve, row].Value;
            Form1.BulkSheet[18, row].Value = db.getRecord(Form1.OrganizedSheet[Sleeve, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[18, row].Value.ToString()))
            {
                Form1.BulkSheet[18, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setNeckStyle(int row) {
            Form1.BulkSheet[9, row].Value = Form1.OrganizedSheet[NeckStyle, row].Value;
            Form1.BulkSheet[19, row].Value = db.getRecord(Form1.OrganizedSheet[NeckStyle, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[19, row].Value.ToString()))
            {
                Form1.BulkSheet[19, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }

        private void setLink(int row)
        {
            Form1.BulkSheet[20, row].Value = Form1.OrganizedSheet[Link, row].Value;
        }
        private void setPrice(int row)
        {
            Form1.BulkSheet[21, row].Value = Form1.OrganizedSheet[Price, row].Value;
        }
        private void setQuantity(int row)
        {
            Form1.BulkSheet[22, row].Value = Form1.OrganizedSheet[Quantity, row].Value;
        }

        private void setItemConnection(int row) {
            try
            {
                if (!Form1.OrganizedSheet[Model, row].Value.ToString().Equals(Form1.OrganizedSheet[Model, row - 1].Value.ToString()))
                {
                    ItemConnection++;
                }
            }
            catch { }
            Form1.BulkSheet[23, row].Value = ItemConnection.ToString();
        }
        private string getSizeBulk(string text)
        {
            string value = "";
            if (!text.ToLower().Contains("y")&&!text.ToLower().Contains(" m")) {



                switch (text.ToLower())
                {
                    case "xxxl":
                        value = "3XL";
                        break;
                    case "xxxxl":
                        value = "4XL";
                        break;
                    case "xxxxxl":
                        value = "5XL";
                        break;
                    case "xxxxxxl":
                        value = "6XL";
                        break;
                    case "xxxxxxxl":
                        value = "7XL";
                        break;
                    case "2xl":
                        value = "XXL";
                        break;
                    default:
                        value = text.ToUpper();
                        break;
                }
            }
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
        private string getTopStyle(string text)
        {
            string value = "";
            if (text.Equals("Two Pieces Wear")) {
                value = "Two Pieces Wear";
            }
            else
            {
                switch (text)
                {
                    case "Sweatshirt":
                        value = "Hoodies & Sweatshirts";
                        break;
                    case "Hoody":
                        value = "Hoodies & Sweatshirts";
                        break;
                    case "Pullover":
                        value = "Pullover Tops";
                        break;
                    case "Vest":
                        value = "Fashion Vests";
                        break;
                    default:
                        value = text + "s";
                        break;
                }
            }
            return value;
        }
        private string getMaterial(string text) {
            string value = "";
            Regex regmat = new Regex(@"([Cc]ot\w+)|([Ll]ea\w+)|([Ww]oo\w+)|([Ss]at\w+)|([Ss]ilk)|([Vv]isco\w+)|([Pp]olye\w+)|([Ll]inen)|([Nn]ylon)|([Vv]elvet)");
            MatchCollection matchmat = regmat.Matches(text);
            if (matchmat.Count > 1)
            {
                value = "Mixed Materials";
            }
            else
            {
                value = textInfo.ToTitleCase(matchmat[0].Value);
            }
            return value.Trim();
        }
    }
}
