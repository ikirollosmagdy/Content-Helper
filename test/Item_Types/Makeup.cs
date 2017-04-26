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
   public class Makeup
    {
        int Brand = 1, Model = 2, type = 3, colors = 4, Desc = 5, Link = 6, price = 7, Quantity = 8
            , UnTranslatedCount = 0;

        private void setupTable()
        {
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));
            DataGridViewComboBoxColumn typeColumn = new DataGridViewComboBoxColumn();
            typeColumn.HeaderText = "Type";
            typeColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            typeColumn.FlatStyle = FlatStyle.Popup;
            typeColumn.Items.AddRange("Blotting Paper", "Blusher", "Body Concealer", "Body Foundation", "Body Powder", "Bronzers", "Eye Concealer", "Eyebrow Color & Shaping", "EyeBrows Enhancers", "Eyelash Enhancers", "Eyeliner", "Eyeshadow", "Eyeshadow Palettes", "Face Foundation", "Face Powder", "Face Primers", "Glitter", "Highlighters & Contour", "Lip & Cheek Tints", "Lip Balm", "Lip Base", "Lip Fixation", "Lip Gloss", "Lip Liners", "Lip Plumpers", "Lipstick", "Lipstick Palettes", "Makeup Finishing", "Makeup Remover", "Makeup Set", "Mascara", "Nail Polish");

            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("NO", "No.")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Model", "Model")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(typeColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Color", "Color")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Desc", "Desc")));
          
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
                        Regex regexModel = new Regex(@"(?!\s+)((\w+([-|/| ])?)?\w+([-|/| ])?\d+(\w+|\d+)?([-|/| ])?(\w+|\d+)?)");
                        MatchCollection matchModel = regexModel.Matches(Form1.Sheet[col, row].Value.ToString());


                        if (matchModel.Count > 0)
                        {
                            Form1.OrganizedSheet[Model, row].Value = matchModel[0].Value;
                            break;

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
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Type", "Type")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArTitle", "Title Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArBrand", "Brand Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDes", "Description Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArType", "Type Arabic")));
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
                        setType(i);
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
            Database db = new Database();
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            string title = Form1.OrganizedSheet[Brand, row].Value + " " + Form1.OrganizedSheet[Model, row].Value + " "
                + Form1.OrganizedSheet[type, row].Value + ", " + Form1.OrganizedSheet[colors, row].Value;
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase(title);
            Form1.BulkSheet[4, row].Value = db.getRecord(Form1.OrganizedSheet[type, row].Value.ToString()) + " من " +
                db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()) + " " + Form1.OrganizedSheet[Model, row].Value +
                "، " + db.getRecord(Form1.OrganizedSheet[colors, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[4, row].Value.ToString()))
            {
                Form1.BulkSheet[4, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[4, row].Style.BackColor = Color.Empty;
            }
            Form1.BulkSheet[11, row].Value = db.getEAN(Form1.BulkSheet[0, row].Value.ToString());
        }
        private void setBrand(int row)
        {
            Database db = new Database();
            Form1.BulkSheet[1, row].Value = Form1.OrganizedSheet[Brand, row].Value;
            Form1.BulkSheet[5, row].Value = db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[5, row].Value.ToString()))
            {
                Form1.BulkSheet[5, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[5, row].Style.BackColor = Color.Empty;
            }
        }
        private void setDescription(int row)
        {
            Database db = new Database();
            string des = "", ArDes = "";
            try
            {

                string[] lines = Form1.OrganizedSheet[Desc, row].Value.ToString().Split('\n');
                foreach (string line in lines)
                {
                    des = des + "<p>" + line + "</p>";
                    ArDes = ArDes + "<p>" + db.getRecord(line) + "</p>";
                }
            }
            catch { }
            des = des + "<p><b>Product Features:</b></p><ul><li>Brand: " + Form1.OrganizedSheet[Brand, row].Value + "</li><li>Type: " +
                Form1.OrganizedSheet[type, row].Value + "</li><li>Color: " + Form1.OrganizedSheet[colors, row].Value + "</li></ul>";
            ArDes = ArDes + "<p><b>خصائص المنتج:</b></p><ul><li>العلامة التجارية: "+db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString())
                + "</li><li>النوع: " +db.getRecord(Form1.OrganizedSheet[type, row].Value.ToString())+ "</li><li>اللون: "+
                db.getRecord(Form1.OrganizedSheet[colors, row].Value.ToString())+ "</li></ul>";
            Form1.BulkSheet[2, row].Value = des;
            Form1.BulkSheet[6, row].Value = ArDes;
            if (CheckEnglish(ArDes))
            {
                Form1.BulkSheet[6, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[6, row].Style.BackColor = Color.Empty;
            }
        }
        private void setType(int row)
        {
            Database db = new Database();
            Form1.BulkSheet[3, row].Value = Form1.OrganizedSheet[type, row].Value;
            Form1.BulkSheet[7, row].Value = db.getRecord(Form1.OrganizedSheet[type, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[7, row].Value.ToString()))
            {
                Form1.BulkSheet[7, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[7, row].Style.BackColor = Color.Empty;
            }

        }
        private void setLink(int row)
        {
            Form1.BulkSheet[8, row].Value = Form1.OrganizedSheet[Link, row].Value;
        }
        private void setPrice(int row)
        {
            Form1.BulkSheet[9, row].Value = Form1.OrganizedSheet[price, row].Value;
        }
        private void setQuantity(int row)
        {
            Form1.BulkSheet[10, row].Value = Form1.OrganizedSheet[Quantity, row].Value;
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
    }
}
