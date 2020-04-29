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
    public class Hand_Tool
    {
        Database db = new Database();
        Common_Use common = new Common_Use();
        TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
        int Brand = 1, Model = 2, Type = 3, Extra = 4, Link = 5, Price = 6, Quantity = 7, UnTranslatedCount = 0;

        private void setupTable()
        {
            DataGridViewComboBoxColumn typeColumn = new DataGridViewComboBoxColumn();
            typeColumn.HeaderText = "Type";
            typeColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            typeColumn.FlatStyle = FlatStyle.Popup;
            typeColumn.Items.AddRange("Axe", "Bench Vise", "Brush", "Chisel", "Clamp", "Cutter", "Cutting Tool", "Engraving Tool", "Filter Key", "Foot Air Pump", "Grease Gun", "Grinding Tool", "Gripper", "Hammer", "Hand Air Pump", "Hand Jack", "Hand Saw", "Hand Tools Organizer", "Leather Hole Punch", "Magnet", "Multi Tool", "Packaging Tool", "Plier", "Polishing Tool", "Reaching Aid", "Riveter", "Sanding Tool", "SandPaper", "Scissor", "Screwdriver", "Shovel ", "Sprayer", "Stapler", "Tester Screwdriver", "Tool Sets", "Turning Tool", "Water Pump", "Wire Stripper", "Wrench");

            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("NO", "No.")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Model", "Model")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(typeColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Extra", "More Data")));

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
                        Regex regexType = new Regex(@"(Stapler|Sprayer|axe|Brush|Chisel|Punch|Clamp|Vice|Cutter|Gripper|Hammer|Plier|Riveter|SandPaper|Scissor|Screwdriver|Wrench)", RegexOptions.IgnoreCase);
                        Match matchType = regexType.Match(Form1.Sheet[col, row].Value.ToString());
                        if (matchType.Success)
                        {
                            Form1.OrganizedSheet[Type, row].Value = common.getReplacement(matchType.Value);
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
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase(string.Format("{0} {1} {2}",
              Form1.OrganizedSheet[Brand, row].Value, Form1.OrganizedSheet[Model, row].Value, Form1.OrganizedSheet[Type, row].Value));

            Form1.BulkSheet[4, row].Value = string.Format("{0} من {1} {2}", db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString()),
                db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()), Form1.OrganizedSheet[Model, row].Value);
            if (common.CheckEnglish(Form1.BulkSheet[4, row].Value.ToString()))
            {
                Form1.BulkSheet[4, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setBrand(int row)
        {
            Form1.BulkSheet[1, row].Value = Form1.OrganizedSheet[Brand, row].Value;
            Form1.BulkSheet[5, row].Value = db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[5, row].Value.ToString()))
            {
                Form1.BulkSheet[5, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }

        }
        private void setDescription(int row)
        {

            string des = "<ul>", ArDes = "<ul>";
            try
            {

                string[] lines = Form1.OrganizedSheet[Extra, row].Value.ToString().Split('\n');
                foreach (string line in lines)
                {
                    des = des + "<li>" + line + "</li>";
                    ArDes = ArDes + "<li>" + db.getRecord(line) + "</li>";
                }
            }
            catch { }
            des = des + "<li>Brand: " + Form1.OrganizedSheet[Brand, row].Value + "</li><li>Type: " + Form1.OrganizedSheet[Type, row].Value + "</li><li>Model: " + Form1.OrganizedSheet[Model, row].Value + "</li></ul>";
            ArDes = ArDes + "<li>العلامة التجارية: " + db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()) + "</li><li>النوع: " + db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString()) +
                "</li><li>الموديل: " + Form1.OrganizedSheet[Model, row].Value + "</li></ul>";

            Form1.BulkSheet[2, row].Value = des;
            Form1.BulkSheet[6, row].Value = ArDes;
            if (common.CheckEnglish(ArDes))
            {
                Form1.BulkSheet[6, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }

        }
        private void setType(int row)
        {
            Form1.BulkSheet[3, row].Value = getAttribType(Form1.OrganizedSheet[Type, row].Value.ToString());
            Form1.BulkSheet[7, row].Value = db.getRecord(Form1.BulkSheet[3, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[7, row].Value.ToString()))
            {
                Form1.BulkSheet[7, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }

        }
        private void setLink(int row)
        {
            Form1.BulkSheet[8, row].Value = Form1.OrganizedSheet[Link, row].Value;
        }
        private void setPrice(int row)
        {
            Form1.BulkSheet[9, row].Value = Form1.OrganizedSheet[Price, row].Value;
        }
        private void setQuantity(int row)
        {
            Form1.BulkSheet[10, row].Value = Form1.OrganizedSheet[Quantity, row].Value;
        }

        private string getAttribType(string text)
        {
            string value = "";
            if (text.Equals("Axe") || text.Equals("Filter Key") || text.Equals("Foot Air Pump") || text.Equals("Hand Air Pump")
                || text.Equals("Magnet") || text.Equals("Riveter") || text.Equals("SandPaper") || text.Equals("Tester Screwdriver")
                || text.Equals("Water Pump") || text.Equals("Tool Sets")

                )
            {
                value = text;
            }
            else
            {
                switch (text)
                {
                    case "Chisel":
                        value = "Chisels & Punches";
                        break;
                    case "Punch":
                        value = "Chisels & Punches";
                        break;
                    case "Clamp":
                        value = "Clamps & Vices";
                        break;
                    case "Vice":
                        value = "Clamps & Vices";
                        break;
                    case "Shovel":
                        value = "Shovel & Trowel";
                        break;
                    case "Trowel":
                        value = "Shovel & Trowel";
                        break;
                    case "Wrench":
                        value = text + "es";
                        break;
                    case "Leather Hole Punch":
                        value = text + "es";
                        break;
                    default:
                        value = text + "s";
                        break;

                }
            }
            return value;

        }

    }
}
