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
    public class Mobile_Accessories
    {
        Database db = new Database();
        int Brand = 1, Model = 2, Type = 3, Mobile = 4, MobileModel = 5, Colors = 6, Material = 7,
            Link = 8, Price = 9, Quantity = 10, UnTranslatedCount = 0;
        private void setupTable()
        {
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("NO", "No.")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Model", "Model")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Type", "Type")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Mobile", "Mobile")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("MobileMO", "Mobile Model")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Color", "Color")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Material", "Material")));

            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Link", "Link")));
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
                        for (int y = 0; y < Form1.Sheet.ColumnCount; y++)
                        {
                            Regex regexModel = new Regex(@"(?!\s+)((\w+([-|/| ])?)?\w+([-|/| ])?\d+(\w+|\d+)?([-|/| ])?(\w+|\d+)?)");
                            MatchCollection matchModel = regexModel.Matches(Form1.Sheet[y, row].Value.ToString());


                            if (matchModel.Count > 0)
                            {
                                Form1.OrganizedSheet[Model, row].Value = matchModel[0].Value;
                                break;
                            }
                        }
                    }
                    catch { }
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
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Color", "Color")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Type", "Type")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Device", "Device")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ColorFa", "Family Color")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Device", "Device Model")));

            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArTitle", "Title Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArBrand", "Brand Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDes", "Description Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArColor", "Color Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArType", "Type Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDevice", "Device Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArColorFa", "Familey Color Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDeviceModel", "Device Model Arabic")));
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
                        setType(i);
                        setDevice(i);
                        setFamilyColor(i);
                        setMobileModel(i);
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
            title = title + Form1.OrganizedSheet[Type, row].Value + " for " + Form1.OrganizedSheet[Mobile, row].Value + " " +
                Form1.OrganizedSheet[MobileModel, row].Value + ", " + Form1.OrganizedSheet[Colors, row].Value;
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase(title);
            Form1.BulkSheet[8, row].Value = db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString()) + " ل" + db.getRecord(Form1.OrganizedSheet[Mobile, row].Value.ToString()) +
                " " + db.getRecord(Form1.OrganizedSheet[MobileModel, row].Value.ToString()) + " " + ArTitle + "، " + db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[8, row].Value.ToString()))
            {
                Form1.BulkSheet[8, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[8, row].Style.BackColor = Color.Empty;
            }
            // Form1.BulkSheet[11, row].Value = db.getEAN(Form1.BulkSheet[0, row].Value.ToString());
        }
        private void setBrand(int row)
        {
            Form1.BulkSheet[1, row].Value = Form1.OrganizedSheet[Brand, row].Value;
            Form1.BulkSheet[9, row].Value = db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[9, row].Value.ToString()))
            {
                Form1.BulkSheet[9, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[9, row].Style.BackColor = Color.Empty;
            }
        }
        private void setDescription(int row)
        {
            Form1.BulkSheet[2, row].Value = "<li>Brand: " + Form1.OrganizedSheet[Brand, row].Value + "</li><li>Mobile Accessory Type: " +
                Form1.OrganizedSheet[Type, row].Value + "</li><li>Compatible with: " + Form1.OrganizedSheet[Mobile, row].Value + " " +
                Form1.OrganizedSheet[MobileModel, row].Value + "</li><li>Color: " + Form1.OrganizedSheet[Colors, row].Value +
                "</li><li>Material: " + Form1.OrganizedSheet[Material, row].Value + "</li><li>High Quality Product</li>";
            Form1.BulkSheet[10, row].Value = "<li>العلامه التجارية: " + db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()) + "</li><li>نوع المنتج: " +
               db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString()) + "</li><li>متوافق مع: " + db.getRecord(Form1.OrganizedSheet[Mobile, row].Value.ToString()) + " " +
               db.getRecord(Form1.OrganizedSheet[MobileModel, row].Value.ToString()) + "</li><li>اللون: " + db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString()) +
                "</li><li>الخامة: " + db.getRecord(Form1.OrganizedSheet[Material, row].Value.ToString()) + "</li><li>منتج عالي الجودة</li>";
            if (CheckEnglish(Form1.BulkSheet[10, row].Value.ToString()))
            {
                Form1.BulkSheet[10, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[10, row].Style.BackColor = Color.Empty;
            }
        }
        private void setColor(int row)
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            Form1.BulkSheet[3, row].Value = textInfo.ToTitleCase(Form1.OrganizedSheet[Colors, row].Value.ToString());
            Form1.BulkSheet[11, row].Value = db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[11, row].Value.ToString()))
            {
                Form1.BulkSheet[11, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[11, row].Style.BackColor = Color.Empty;
            }
        }
        private void setType(int row)
        {
            Form1.BulkSheet[4, row].Value = getTypeAttribute(Form1.OrganizedSheet[Type, row]);
            Form1.BulkSheet[12, row].Value = db.getRecord(Form1.BulkSheet[4, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[12, row].Value.ToString()))
            {
                Form1.BulkSheet[12, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[12, row].Style.BackColor = Color.Empty;
            }
        }
        private void setDevice(int row)
        {
            Form1.BulkSheet[5, row].Value = Form1.OrganizedSheet[Mobile, row].Value;
            Form1.BulkSheet[13, row].Value = db.getRecord(Form1.OrganizedSheet[Mobile, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[13, row].Value.ToString()))
            {
                Form1.BulkSheet[13, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[13, row].Style.BackColor = Color.Empty;
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
                    Form1.BulkSheet[6, row].Value = getReplacement(value);
                }
                else
                {
                    Form1.BulkSheet[6, row].Value = "Multi Color";
                }
            }
            else
            {
                Form1.BulkSheet[6, row].Value = getReplacement(Form1.OrganizedSheet[Colors, row].Value.ToString());
            }

            Form1.BulkSheet[14, row].Value = db.getRecord(Form1.BulkSheet[6, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[14, row].Value.ToString()))
            {
                Form1.BulkSheet[14, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[14, row].Style.BackColor = Color.Empty;
            }

        }
        private void setMobileModel(int row)
        {
            Form1.BulkSheet[7, row].Value = Form1.OrganizedSheet[MobileModel, row].Value;
            Form1.BulkSheet[15, row].Value = db.getRecord(Form1.OrganizedSheet[MobileModel, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[15, row].Value.ToString()))
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
        bool CheckEnglish(string text)
        {
            bool IsEnglish = false;
            Regex regex = new Regex(@"[^pulbi<>\/\d\.,\s]([a-zA-Z])");
            Match match = regex.Match(text);
            if (match.Success)
            {
                IsEnglish = true;
            }
            else
            {
                IsEnglish = false;
            }

            return IsEnglish;
        }

        private string getTypeAttribute(DataGridViewCell cell)
        {
            string result = "";
            if (cell.Value.ToString().ToLower().Contains("cover"))
            {
                result = "Cases & Covers";
            }
            else if (cell.Value.ToString().ToLower().Contains("holder"))
            {
                result = "Mounts & Holders";
            }
            else if (cell.Value.ToString().ToLower().Contains("selfie"))
            {
                result = "Selfie Stick";
            }

            return result;
        }
        string getReplacement(String Text)
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            string arabicMatch = Text;
            if (Text != null)
            {
                Text = Text.ToLower().Trim();
                try
                {
                    foreach (string line in System.IO.File.ReadAllLines("lookup.dat"))
                    {

                        if (line.Contains(Text))
                            arabicMatch = line.Split('	')[1];

                    }
                }
                catch (Exception)
                {
                    arabicMatch = "";
                }
            }
            return textInfo.ToTitleCase(arabicMatch);
        }
    }
}
