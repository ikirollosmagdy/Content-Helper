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
    public class Rings
    {
        Database db = new Database();
        TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
        Common_Use common = new Common_Use();
        int Brand = 1, Model = 2, Type = 4, Gender = 3, Material = 5, Size = 6, Stone_Type = 7, Style = 8,
           Link = 9, Price = 10, Quantity = 11, UnTranslatedCount = 0;
        private void setupTable()
        {

            DataGridViewComboBoxColumn TypeColumn = new DataGridViewComboBoxColumn();
            TypeColumn.HeaderText = "Type";
            TypeColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            TypeColumn.FlatStyle = FlatStyle.Popup;
            TypeColumn.Items.AddRange("3-Stone Diamond", "Band", "Couple", "Dangle", "Fashion", "Midi", "Multi Finger", "Nail", "Palm Cuff", "Pearl", "Promise", "Smart", "Stacking", "Toe");

            DataGridViewComboBoxColumn MaterialColumn = new DataGridViewComboBoxColumn();
            MaterialColumn.HeaderText = "Material";
            MaterialColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            MaterialColumn.FlatStyle = FlatStyle.Popup;
            MaterialColumn.Items.AddRange("Alloy", "Beads", "Brass", "Bronze", "Ceramic", "Copper", "Gold", "Gold Plated", "Leather", "Mixed", "Palladium", "Plastic", "Platinum", "Platinum Plated Stainless Steel", "Rhodium Plated", "Rope", "Rose Gold", "Silver", "Silver Plated", "Stainless Steel", "String", "Three Tone Gold", "Titanium", "Tungsten", "Two Tone Gold", "White Gold", "White Gold Plated");

            DataGridViewComboBoxColumn StoneTypeColumn = new DataGridViewComboBoxColumn();
            StoneTypeColumn.HeaderText = "Stone type";
            StoneTypeColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            StoneTypeColumn.FlatStyle = FlatStyle.Popup;
            StoneTypeColumn.Items.AddRange("Agate", "Amber", "Amethyst", "Astrophyllite", "Aveturine", "Beads", "Citrine", "Coral", "Crystals", "Cubic Zirconia", "Diamond", "Emerald", "Enamel", "Garnet", "Glass", "Imitation Pearl", "Jade", "Jasper", "Lapis Lazuli", "Larimar", "Moonstone", "Morganite", "Multi Stones", "Natural Pearl", "Obsidian", "Onyx", "Opal", "Peridot", "Porcelain", "Quartz", "Rhinestone", "Rhodonite", "Ruby", "Sandstone", "Sapphire", "Shells", "Swarovski", "Tanzanite", "Tiger's Eye Stone", "Topaz", "Turquoise", "Unakite", "Without Stones");

            DataGridViewComboBoxColumn StyleColumn = new DataGridViewComboBoxColumn();
            StyleColumn.HeaderText = "Style";
            StyleColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            StyleColumn.FlatStyle = FlatStyle.Popup;
            StyleColumn.Items.AddRange("Casual", "Religious", "Wedding & Engagement");

            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("NO", "No.")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Model", "Model")));

            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Gender", "Gender")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(TypeColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(MaterialColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("size", "Size")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(StoneTypeColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(StyleColumn)));
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
                        if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("type"))
                        {
                            Form1.OrganizedSheet[Type, row].Value = Form1.Sheet[col, row].Value;

                        }
                        else
                        {
                            Regex regexType = new Regex(@"(3-Stone Diamond|Band|Couple|Dangle|Fashion|Midi|Multi Finger|Nail|Palm Cuff|Pearl|Promise|Smart|Stacking|Toe)");
                            Match matchType = regexType.Match(Form1.Sheet[col, row].Value.ToString());
                            if (matchType.Success)
                            {
                                Form1.OrganizedSheet[Type, row].Value = common.getReplacement(matchType.Value);

                            }
                        }
                    }
                    catch { }
                    try
                    {
                        if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("gender"))
                        {
                            Form1.OrganizedSheet[Gender, row].Value = Form1.Sheet[col, row].Value;

                        }


                    }
                    catch (Exception) { }
                    try
                    {
                        if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("material"))
                        {
                            Form1.OrganizedSheet[Material, row].Value = Form1.Sheet[col, row].Value;

                        }


                    }
                    catch (Exception) { }
                    try
                    {
                        if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("size"))
                        {
                            Form1.OrganizedSheet[Size, row].Value = Form1.Sheet[col, row].Value;

                        }


                    }
                    catch (Exception) { }
                    try
                    {
                        if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("stone"))
                        {
                            Form1.OrganizedSheet[Stone_Type, row].Value = Form1.Sheet[col, row].Value.ToString().Trim();
                        }

                    }
                    catch { }
                    try
                    {
                        if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("style"))
                        {
                            Form1.OrganizedSheet[Style, row].Value = Form1.Sheet[col, row].Value.ToString().Trim();
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
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("gender", "Gender")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("type", "Type")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("material", "Material")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("size", "Size")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("stone", "Stone Type")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("style", "Ring Style")));

            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArTitle", "Title Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArBran", "Brand Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDes", "Description Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Argender", "Gender Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Artype", "Type Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Armaterial", "Material Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arsize", "Size Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arstone", "Stone Type Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arstyle", "Ring Style Arabic")));

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
                        setGender(i);
                        setType(i);
                        setMaterial(i);
                        setSize(i);
                        setStoneType(i);
                        setStyle(i);
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
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase(string.Format("{0} {1} {2} ring for {3} - {4}",
               Form1.OrganizedSheet[Brand, row].Value, Form1.OrganizedSheet[Model, row].Value, Form1.OrganizedSheet[Type, row].Value,
               Form1.OrganizedSheet[Gender, row].Value, Form1.OrganizedSheet[Size, row].Value));
            Form1.BulkSheet[9, row].Value = string.Format("خاتم {0} لل{1} من {2} {3} - {4}", db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString()),
                db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()),
                Form1.OrganizedSheet[Model, row].Value, db.getRecord(Form1.OrganizedSheet[Size, row].Value.ToString()));
            if (common.CheckEnglish(Form1.BulkSheet[9, row].Value.ToString()))
            {
                Form1.BulkSheet[9, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setBrand(int row)
        {
            Form1.BulkSheet[1, row].Value = Form1.OrganizedSheet[Brand, row].Value;
            Form1.BulkSheet[10, row].Value = db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[10, row].Value.ToString()))
            {
                Form1.BulkSheet[10, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setDescription(int row)
        {
            Form1.BulkSheet[2, row].Value = string.Format("<p><b>Product Features:</b></p><ul><li>Brand:{0}</li><li>Model Number:{1}</li><li>Stone Type:{2}</li><li>Size:{3}</li><li>Material:{4}</li><li>Targeted Group:{5}</li></ul>",
                Form1.OrganizedSheet[Brand, row].Value, Form1.OrganizedSheet[Model, row].Value, Form1.OrganizedSheet[Stone_Type, row].Value, Form1.OrganizedSheet[Size, row].Value, Form1.OrganizedSheet[Material, row].Value, Form1.OrganizedSheet[Gender, row].Value);
            Form1.BulkSheet[11, row].Value = string.Format("<p><b>خصائص المنتج:</b></p><ul><li>العلامة  التجارية:{0}</li><li>رقم الموديلr:{1}</li><li>نوع الحجر:{2}</li><li>المقاس:{3}</li><li>الخامة:{4}</li><li>المجموعة المستهدفة:{5}</li></ul>",
               db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()), Form1.OrganizedSheet[Model, row].Value, db.getRecord(Form1.OrganizedSheet[Stone_Type, row].Value.ToString()), Form1.OrganizedSheet[Size, row].Value, db.getRecord(Form1.OrganizedSheet[Material, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString()));
            if (common.CheckEnglish(Form1.BulkSheet[11, row].Value.ToString()))
            {
                Form1.BulkSheet[11, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setGender(int row)
        {
            Form1.BulkSheet[3, row].Value = textInfo.ToTitleCase(Form1.OrganizedSheet[Gender, row].Value.ToString().Trim());
            Form1.BulkSheet[12, row].Value = db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[12, row].Value.ToString()))
            {
                Form1.BulkSheet[12, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setType(int row)
        {
            Form1.BulkSheet[4, row].Value = getType(Form1.OrganizedSheet[Type, row].Value.ToString());
            Form1.BulkSheet[13, row].Value = db.getRecord(Form1.BulkSheet[4, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[13, row].Value.ToString()))
            {
                Form1.BulkSheet[13, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setMaterial(int row)
        {
            Form1.BulkSheet[5, row].Value = Form1.OrganizedSheet[Material, row].Value;
            Form1.BulkSheet[14, row].Value = db.getRecord(Form1.BulkSheet[5, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[14, row].Value.ToString()))
            {
                Form1.BulkSheet[14, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }

        }
        private void setSize(int row)
        {
            Form1.BulkSheet[6, row].Value = Form1.OrganizedSheet[Size, row].Value;
            Form1.BulkSheet[15, row].Value = db.getRecord(Form1.OrganizedSheet[Size, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[15, row].Value.ToString()))
            {
                Form1.BulkSheet[15, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setStoneType(int row)
        {
            Form1.BulkSheet[7, row].Value = textInfo.ToTitleCase(Form1.OrganizedSheet[Stone_Type, row].Value.ToString().Trim());
            Form1.BulkSheet[16, row].Value = db.getRecord(Form1.OrganizedSheet[Stone_Type, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[16, row].Value.ToString()))
            {
                Form1.BulkSheet[16, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setStyle(int row)
        {
            Form1.BulkSheet[8, row].Value = Form1.OrganizedSheet[Style, row].Value;
            Form1.BulkSheet[17, row].Value = db.getRecord(Form1.OrganizedSheet[Style, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[17, row].Value.ToString()))
            {
                Form1.BulkSheet[17, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }


        private void setLink(int row)
        {
            Form1.BulkSheet[18, row].Value = Form1.OrganizedSheet[Link, row].Value;
        }
        private void setPrice(int row)
        {
            Form1.BulkSheet[19, row].Value = Form1.OrganizedSheet[Price, row].Value;
        }
        private void setQuantity(int row)
        {
            Form1.BulkSheet[20, row].Value = Form1.OrganizedSheet[Quantity, row].Value;
        }



        private string getType(string text)
        {

            switch (text)
            {
                case "Couple":
                case "Fashion":
                case "Midi":
                case "Nail":
                    return text + " Rings";

                default:
                    return text;
            }
        }
    }
}
