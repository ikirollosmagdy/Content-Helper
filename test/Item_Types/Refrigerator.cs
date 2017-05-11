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
    public class Refrigerator


    {
        int Brand = 1, Model = 2, Type = 3, Capacity = 4, Colors = 5, Des = 6, Dims = 7, Material = 8
           , style = 9, Shipping=10, Link = 11, Price = 12, Quantity = 13, UnTranslatedCount=0;
        Database db = new Database();
        Common_Use common = new Common_Use();

        private void setupTable()
        {
            DataGridViewComboBoxColumn MaterialColumn = new DataGridViewComboBoxColumn();
            MaterialColumn.HeaderText = "Material";
            MaterialColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            MaterialColumn.FlatStyle = FlatStyle.Popup;
            MaterialColumn.Items.AddRange("Metal", "Plastic", "Stainless Platinum", "Stainless Steel");

            DataGridViewComboBoxColumn styleColumn = new DataGridViewComboBoxColumn();
            styleColumn.HeaderText = "Style";
            styleColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            styleColumn.FlatStyle = FlatStyle.Popup;
            styleColumn.Items.AddRange("Chest Freezer", "Compact", "Freezer", "Freezer on Bottom", "Freezer on Top", "Freezerless", "French Door", "Side by Side", "Upright Freezer"
);

            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("No", "Row Number")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Model", "Model")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Type", "Type")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Capacity", "Capacity")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Color", "Color")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Des", "Desciption")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Dims", "Dimensions(Depth x Height x Width")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(MaterialColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(styleColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Shipping", "Shipping weight")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Link", "Link")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Price", "Price")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Quantity", "Quantity")));
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
                        Regex regexSize = new Regex(@"([Rr]efr?\w+)|(\w+eezer)");
                        Match matchSize = regexSize.Match(Form1.Sheet[col, row].Value.ToString());
                        if (matchSize.Success)
                        {
                            Form1.OrganizedSheet[Type, row].Value = common.getReplacement(matchSize.Value);

                         }
                    }
                    catch { }
                    try
                    {
                        for (int y = 0; y < Form1.Sheet.ColumnCount; y++)
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
                        if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("color"))
                        {
                            Form1.OrganizedSheet[Colors, row].Value = Form1.Sheet[col, row].Value.ToString().Trim();
                        }
                        else
                        {
                            Regex regexColor = new Regex(@"([Bb]eige|[Bb]lack|[Bb]lue|[Bb]rown|[Cc]lear|[Gg]old|[Gg]reen|[Gg]rey|[Mm]ultiColor|[Oo]ffWhite|[Oo]range|[Pp]ink|[Pp]urple|[Rr]ed|[Ss]ilver|[Tt]urquoise|[Ww]hite|[Yy]ellow)");
                            MatchCollection matchColor = regexColor.Matches(Form1.Sheet[col, row].Value.ToString());


                            if (matchColor.Count > 0)
                            {
                                Form1.OrganizedSheet[Colors, row].Value = matchColor[0].Value;


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

        public void creatBulk()
        {
            UnTranslatedCount = 0;
            Form1.txtUntranslated.GetCurrentParent().Invoke(new Action(() => Form1.txtUntranslated.Text = UnTranslatedCount.ToString()));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Clear()));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Title", "Title")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Bran", "Brand")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Des", "Description")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Color", "Color")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Depth", "Depth")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Height", "Height")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Width", "Width")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Material", "Material")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("model", "Model Number")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Capacity", "Capacity")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Style", "Style")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Weight", "Shipping Weight")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArTitle", "Title Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArBrand", "Brand Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDes", "Description Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArColor", "Color Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDepth", "Depth Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArHeight", "Height Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArWidth", "Width Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArMaterial", "Material Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArModel", "Model Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArStyle", "Style Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArShip", "Shipping Weight Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Link", "Link")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Price", "Price")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Quantity", "Quantity")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("EAn", "Suggested EAN")));
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
                        setDimensions(i);
                        SetMaterial(i);
                        setModel(i);
                        setCapacity(i);
                        setStyle(i);
                        setShipping(i);
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



        void setTitle(int row)
        {
            Database db = new Database();
            
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            string title = Form1.OrganizedSheet[Brand, row].Value + " " + Form1.OrganizedSheet[Model, row].Value + " " + Form1.OrganizedSheet[Type, row].Value + " - "
                + Form1.OrganizedSheet[Capacity, row].Value + ", " + Form1.OrganizedSheet[Colors, row].Value;
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase(title);
            Form1.BulkSheet[12, row].Value = db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString()) + " " +
                db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()) + " " + db.getRecord(Form1.OrganizedSheet[Capacity, row].Value.ToString())
                + " " + Form1.OrganizedSheet[Model, row].Value + " - " + db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[12, row].Value.ToString()))
            {
                Form1.BulkSheet[12, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[12, row].Style.BackColor = Color.Empty;
            }
            Form1.BulkSheet[26, row].Value = db.getEAN(Form1.BulkSheet[0, row].Value.ToString());

        }
        void setBrand(int row)
        {
            Database db = new Database();
            Form1.BulkSheet[1, row].Value = Form1.OrganizedSheet[Brand, row].Value;
            Form1.BulkSheet[13, row].Value = db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[13, row].Value.ToString()))
            {
                Form1.BulkSheet[13, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[13, row].Style.BackColor = Color.Empty;
            }
        }
        void setColor(int row)
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            Database db = new Database();
            Form1.BulkSheet[3, row].Value = textInfo.ToTitleCase( Form1.OrganizedSheet[Colors, row].Value.ToString());
            Form1.BulkSheet[15, row].Value = db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[15, row].Value.ToString()))
            {
                Form1.BulkSheet[15, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[15, row].Style.BackColor = Color.Empty;
            }
        }
        void setModel(int row)
        {
            Form1.BulkSheet[8, row].Value = Form1.OrganizedSheet[Model, row].Value;
            Form1.BulkSheet[20, row].Value = Form1.OrganizedSheet[Model, row].Value;
        }

        void setCapacity(int row)
        {
           
            Form1.BulkSheet[9, row].Value = Form1.OrganizedSheet[Capacity, row].Value;
            
        }

        void SetMaterial(int row)
        {
            Database db = new Database();
            Form1.BulkSheet[7, row].Value = Form1.OrganizedSheet[Material, row].Value;
            Form1.BulkSheet[19, row].Value = db.getRecord(Form1.OrganizedSheet[Material, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[19, row].Value.ToString()))
            {
                Form1.BulkSheet[19, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[19, row].Style.BackColor = Color.Empty;
            }
        }
        void setDescription(int row)
        {
            Database db = new Database();
            string des = "",ArDes="";
            try
            {

                string[] lines = Form1.OrganizedSheet[Des, row].Value.ToString().Split('\n');
                foreach (string line in lines)
                {
                    des = des + "<p>" + line + "</p>";
                    ArDes = ArDes + "<p>" + db.getRecord(line) + "</p>";
                }
            }
            catch { }
            des = des + "<ul><li>Color: " + Form1.OrganizedSheet[Colors, row].Value + "</li><li>Dimensions: " + Form1.OrganizedSheet[Dims, row].Value +
"</li><li>Capacity: " + Form1.OrganizedSheet[Capacity, row].Value + "</li><li>Style: " + Form1.OrganizedSheet[style, row].Value + "</li></ul>";
            ArDes = ArDes + "<ul><li>اللون: " + db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString()) + "</li><li>الابعاد: " +
                db.getRecord(Form1.OrganizedSheet[Dims, row].Value.ToString()) + "</li><li>السعة: " + db.getRecord(Form1.OrganizedSheet[Capacity, row].Value.ToString()) +
                "</li><li>الشكل: " + db.getRecord(Form1.OrganizedSheet[style, row].Value.ToString()) + "</li></ul>";

            Form1.BulkSheet[2, row].Value = des;
            Form1.BulkSheet[14, row].Value = ArDes;
            if (common.CheckEnglish(ArDes))
            {
                Form1.BulkSheet[14, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[14, row].Style.BackColor = Color.Empty;
            }

        }
        void setDimensions(int row)
        {
            if (Form1.OrganizedSheet[Dims, row].Value.ToString().Contains("x"))
            {
                try
                {
                    string[] dim = Form1.OrganizedSheet[Dims, row].Value.ToString().Split('x');
                    Form1.BulkSheet[4, row].Value = dim[0];
                    Form1.BulkSheet[5, row].Value = dim[1];
                    Form1.BulkSheet[6, row].Value = dim[2];
                    Form1.BulkSheet[16, row].Value = dim[0];
                    Form1.BulkSheet[17, row].Value = dim[1];
                    Form1.BulkSheet[18, row].Value = dim[2];
                }
                catch { }

            }
            else if (Form1.OrganizedSheet[Dims, row].Value.ToString().Contains("*"))
            {
                try
                {
                    string[] dim = Form1.OrganizedSheet[Dims, row].Value.ToString().Split('*');
                    Form1.BulkSheet[4, row].Value = dim[0];
                    Form1.BulkSheet[5, row].Value = dim[1];
                    Form1.BulkSheet[6, row].Value = dim[2];
                    Form1.BulkSheet[16, row].Value = dim[0];
                    Form1.BulkSheet[17, row].Value = dim[1];
                    Form1.BulkSheet[18, row].Value = dim[2];
                }
                catch { }
            }
            else if (Form1.OrganizedSheet[Dims, row].Value.ToString().Contains("-"))
            {
                try
                {
                    string[] dim = Form1.OrganizedSheet[Dims, row].Value.ToString().Split('-');
                    Form1.BulkSheet[4, row].Value = dim[0];
                    Form1.BulkSheet[5, row].Value = dim[1];
                    Form1.BulkSheet[6, row].Value = dim[2];
                    Form1.BulkSheet[16, row].Value = dim[0];
                    Form1.BulkSheet[17, row].Value = dim[1];
                    Form1.BulkSheet[18, row].Value = dim[2];
                }
                catch { }
            }

        }
        private void setStyle(int row)
        {
            Database db = new Database();
            Form1.BulkSheet[10, row].Value = Form1.OrganizedSheet[style, row].Value;
            Form1.BulkSheet[21, row].Value = db.getRecord(Form1.OrganizedSheet[style, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[19, row].Value.ToString()))
            {
                Form1.BulkSheet[19, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[19, row].Style.BackColor = Color.Empty;
            }
        }

        private void setShipping(int row)
        {
            Form1.BulkSheet[11, row].Value = Form1.OrganizedSheet[Shipping, row].Value;
            Form1.BulkSheet[22, row].Value = Form1.OrganizedSheet[Shipping, row].Value;
        }
        private void setLink(int row)
        {
            Form1.BulkSheet[23, row].Value = Form1.OrganizedSheet[Link, row].Value;
        }
        private void setPrice(int row)
        {
            Form1.BulkSheet[24, row].Value = Form1.OrganizedSheet[Price, row].Value;
        }
        private void setQuantity(int row)
        {
            Form1.BulkSheet[25, row].Value = Form1.OrganizedSheet[Quantity, row].Value;
        }

      
    }
}
