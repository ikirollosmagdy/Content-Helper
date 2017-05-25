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
  public  class Watches
    {
        Database db = new Database();
        Common_Use common = new Common_Use();
        int Brand = 1, Model = 2, DisplayType = 3, Type = 4, Gender = 5, WatchShap = 6,BandMaterial=7,
           Link = 8, Price = 9, Quantity = 10, UnTranslatedCount = 0;

        private void setupTable()
        {
            DataGridViewComboBoxColumn DisplaytypeColumn = new DataGridViewComboBoxColumn();
            DisplaytypeColumn.HeaderText = "Display Type";
            DisplaytypeColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            DisplaytypeColumn.FlatStyle = FlatStyle.Popup;
            DisplaytypeColumn.Items.AddRange("Analog", "Analog-Digital", "Digital", "LED");

            DataGridViewComboBoxColumn TypeColumn = new DataGridViewComboBoxColumn();
            TypeColumn.HeaderText = "Type";
            TypeColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            TypeColumn.FlatStyle = FlatStyle.Popup;
            TypeColumn.Items.AddRange("Casual Watch", "Dress Watch", "Fan Sport Watch", "Kids Watch", "Pocket Watch", "Sport Watch", "Watch Set");

            DataGridViewComboBoxColumn GenderColumn = new DataGridViewComboBoxColumn();
            GenderColumn.HeaderText = "Gender";
            GenderColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            GenderColumn.FlatStyle = FlatStyle.Popup;
            GenderColumn.Items.AddRange("Boys", "Girls", "Kids", "Men", "Unisex", "Women");

            DataGridViewComboBoxColumn BandMaterial = new DataGridViewComboBoxColumn();
            BandMaterial.HeaderText = "Band Material";
            BandMaterial.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            BandMaterial.FlatStyle = FlatStyle.Popup;
            BandMaterial.Items.AddRange("18K Rose Gold", "18K White Gold", "18K Yellow Gold", "22K Rose Gold", "22K White Gold", "22K Yellow Gold", "Alloy", "Aluminum", "Ceramic", "Fabric", "Fiber", "Leather", "Metal", "Mixed", "Nylon", "pearl", "Plastic", "Platinum", "Polycarbonate", "Polyurethane", "Resin", "Rubber", "Silicone", "Silver", "Stainless Steel", "Synthetic", "Titanium", "White Gold Plated", "Wood", "Yellow Gold Plated");

            DataGridViewComboBoxColumn watchShape = new DataGridViewComboBoxColumn();
            watchShape.HeaderText = "Watch shape";
            watchShape.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            watchShape.FlatStyle = FlatStyle.Popup;
            watchShape.Items.AddRange("Crown", "Heart", "Hexa", "Inlay", "Letter", "Mixed", "Octagon", "Oval", "Rectangle", "Rhombus", "Round", "Square", "Tonneau", "Triangle", "X-shape");

            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("NO", "No.")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Model", "Model")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(DisplaytypeColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(TypeColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(GenderColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(watchShape)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(BandMaterial)));
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

                        Regex regexGender = new Regex(@"(([Ww]o)?[Mm][ae]n)|([Uu]ni(sex)?)|([Bb]oys?)|([Gg]irls?)");
                        Match matchGender = regexGender.Match(Form1.Sheet.Rows[row].Cells[col].Value.ToString());
                        if (matchGender.Success)
                        {
                            Form1.OrganizedSheet.Rows[row].Cells[Gender].Value = common.getReplacement(matchGender.Value);

                        }
                        

                    }
                    catch (Exception) { }
                    try {
                        Regex regexType = new Regex(@"([Cc]asual)|([Dd]ress)|([Ss]port)|([Kk]id)");
                        Match matchType= regexType.Match(Form1.Sheet.Rows[row].Cells[col].Value.ToString());
                        if (matchType.Success)
                        {
                            Form1.OrganizedSheet[Type,row].Value= common.getReplacement(matchType.Value);
                        }

                    } catch { }

                    try {
                        Regex regexDisplayType = new Regex(@"([Aa]nalog)|([Dd]gt)|([Ll][Ee][Dd])|([Dd]igital)|([Mm]ix)");
                        Match matchDisplayType = regexDisplayType.Match(Form1.Sheet.Rows[row].Cells[col].Value.ToString());
                        if (matchDisplayType.Success)
                        {
                            Form1.OrganizedSheet[DisplayType, row].Value = common.getReplacement(matchDisplayType.Value);
                        }

                    } catch { }


                    try
                    {
                        if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("material"))
                        {
                            Form1.OrganizedSheet[BandMaterial, row].Value = Form1.Sheet[col, row].Value.ToString().Trim();
                        }
                        else
                        {
                            Regex regexMaterial = new Regex(@"([Pp]las\w+)|([Ll]ea\w+)|([Mm]eta\w+)|([Ww]ood)|([Rr]ub\w+)|([Ss]ili\w+)|([Rr]es\w+)");
                            MatchCollection matchMaterial = regexMaterial.Matches(Form1.Sheet[col, row].Value.ToString());


                            if (matchMaterial.Count > 0)
                            {
                                Form1.OrganizedSheet[BandMaterial, row].Value =common.getReplacement( matchMaterial[0].Value.Trim());


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
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("wat", "Watch Shape")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("band", "Band Material")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("DisType", "Display Type")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("gender", "Gender")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("type", "Type")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("model", "Model")));

            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArTitle", "Title Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArBran", "Brand Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDes", "Description Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arwat", "Watch Shape Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arband", "Band Material Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDisType", "Display Type Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Argender", "Gender Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Artype", "Type Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Armodel", "Model Arabic")));

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
                      setWatchShape(i);
                       setBandMaterial(i);
                      setDisplayType(i);
                     setGender(i);
                        setType(i);
                      setModel(i);

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
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase( string.Format("{0} {1} {2} {3} for {4}",Form1.OrganizedSheet[Brand,row].Value
                , Form1.OrganizedSheet[Model, row].Value, Form1.OrganizedSheet[DisplayType, row].Value, Form1.OrganizedSheet[Type, row].Value
                , Form1.OrganizedSheet[Gender, row].Value));
            Form1.BulkSheet[9, row].Value = string.Format("{0} {1} لل{2} من {3} {4}",db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString()),
                db.getRecord(Form1.OrganizedSheet[DisplayType, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString())
               , db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()), Form1.OrganizedSheet[Model, row].Value);
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
        private void setDescription(int row) {
            Form1.BulkSheet[2, row].Value = string.Format("<p><b>Product Features:</b></p><ul><li>Brand:{0}</li><li>Model Number:{1}</li><li>Type:{2}</li><li>Watch Shape:{3}</li><li>Band Material:{4}</li><li>Display Type:{5}</li><li>Targeted Group:{6}</li></ul> ",
                Form1.OrganizedSheet[Brand,row].Value, Form1.OrganizedSheet[Model, row].Value, Form1.OrganizedSheet[Type, row].Value, Form1.OrganizedSheet[WatchShap, row].Value, Form1.OrganizedSheet[BandMaterial, row].Value, Form1.OrganizedSheet[DisplayType, row].Value, Form1.OrganizedSheet[Gender, row].Value);
            Form1.BulkSheet[11, row].Value = string.Format("<p><b>خصائص المنتج:</b></p><ul><li>العلامة التجارية:{0}</li><li>رقم الموديل:{1}</li><li>النوع:{2}</li><li>الشكل:{3}</li><li>خامة السوار:{4}</li><li>نوع العرض:{5}</li><li>المجموعة المستهدفة:{6}</li></ul>",
            db.getRecord( Form1.OrganizedSheet[Brand, row].Value.ToString()), Form1.OrganizedSheet[Model, row].Value, db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString()),
            db.getRecord(Form1.OrganizedSheet[WatchShap, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[BandMaterial, row].Value.ToString()),
           db.getRecord(Form1.OrganizedSheet[DisplayType, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString()));
            if (common.CheckEnglish(Form1.BulkSheet[11, row].Value.ToString()))
            {
                Form1.BulkSheet[11, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setWatchShape(int row)
        {
            Form1.BulkSheet[3, row].Value = Form1.OrganizedSheet[WatchShap, row].Value;
            Form1.BulkSheet[12, row].Value = db.getRecord(Form1.OrganizedSheet[WatchShap, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[12, row].Value.ToString()))
            {
                Form1.BulkSheet[12, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setBandMaterial(int row)
        {
            Form1.BulkSheet[4, row].Value = Form1.OrganizedSheet[BandMaterial, row].Value;
            Form1.BulkSheet[13, row].Value = db.getRecord(Form1.OrganizedSheet[BandMaterial, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[13, row].Value.ToString()))
            {
                Form1.BulkSheet[13, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setDisplayType(int row) {
            Form1.BulkSheet[5, row].Value = Form1.OrganizedSheet[DisplayType, row].Value;
            Form1.BulkSheet[14, row].Value = db.getRecord(Form1.OrganizedSheet[DisplayType, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[14, row].Value.ToString()))
            {
                Form1.BulkSheet[14, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setGender(int row)
        {
            Form1.BulkSheet[6, row].Value = Form1.OrganizedSheet[Gender, row].Value;
            Form1.BulkSheet[15, row].Value = db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[15, row].Value.ToString()))
            {
                Form1.BulkSheet[15, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setType(int row)
        {
            Form1.BulkSheet[7, row].Value = Form1.OrganizedSheet[Type, row].Value;
            Form1.BulkSheet[16, row].Value = db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[16, row].Value.ToString()))
            {
                Form1.BulkSheet[16, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setModel(int row)
        {
            Form1.BulkSheet[8, row].Value = Form1.OrganizedSheet[Model, row].Value;
            Form1.BulkSheet[17, row].Value = Form1.OrganizedSheet[Model, row].Value;
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
    }
}
