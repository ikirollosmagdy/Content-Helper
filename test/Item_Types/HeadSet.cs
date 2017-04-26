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
   public class HeadSet
    {
        Database db = new Database();
        int Brand = 1, Model = 2, Type = 3, Colors = 4, Connectivty = 5, Device = 6,Micophone=7,Noise=8,Extra=9,
            Link = 10, Price = 11, Quantity = 12, UnTranslatedCount = 0;
        private void setupTable()
        {
            DataGridViewComboBoxColumn typeColumn = new DataGridViewComboBoxColumn();
            typeColumn.HeaderText = "Type";
            typeColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            typeColumn.FlatStyle = FlatStyle.Popup;
            typeColumn.Items.AddRange("In Ear", "On Ear", "Over Ear");

            DataGridViewComboBoxColumn Devicecolumn = new DataGridViewComboBoxColumn();
            Devicecolumn.HeaderText = "Compatible with";
            Devicecolumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            Devicecolumn.FlatStyle = FlatStyle.Popup;
            Devicecolumn.Items.AddRange("All", "Mac", "Mobile Phones", "MP3 Players", "MP4 Players", "Tablets");

            DataGridViewComboBoxColumn MicColumn = new DataGridViewComboBoxColumn();
            MicColumn.HeaderText = "Microphone";
            MicColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            MicColumn.FlatStyle = FlatStyle.Popup;
            MicColumn.Items.AddRange("YES","NO");

            DataGridViewComboBoxColumn NoiseCoulmn = new DataGridViewComboBoxColumn();
            NoiseCoulmn.HeaderText = "Noise";
            NoiseCoulmn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            NoiseCoulmn.FlatStyle = FlatStyle.Popup;
            NoiseCoulmn.Items.AddRange("YES", "NO");

            DataGridViewComboBoxColumn ConeectivtyColumn = new DataGridViewComboBoxColumn();
            ConeectivtyColumn.HeaderText = "Connectivity";
            ConeectivtyColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            ConeectivtyColumn.FlatStyle = FlatStyle.Popup;
            ConeectivtyColumn.Items.AddRange("Wired", "Wired/Wireless", "Wireless");

            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("NO", "No.")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Model", "Model")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(typeColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Color", "Color")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(ConeectivtyColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(Devicecolumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(MicColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(NoiseCoulmn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("extra", "Description")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Link", "Link")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Price", "Price")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Qunatity", "Quantity")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.RowCount = Form1.Sheet.RowCount));
            foreach (DataGridViewColumn co in Form1.OrganizedSheet.Columns)
            {
                co.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

        }
        public void Organize() {
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

                }
            }

        }
        public void createBulk() {
            UnTranslatedCount = 0;
            Form1.txtUntranslated.GetCurrentParent().Invoke(new Action(() => Form1.txtUntranslated.Text = UnTranslatedCount.ToString()));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Clear()));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Title", "Title")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Bran", "Brand")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Des", "Description")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("color", "Color")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Model", "Model")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("connectivity", "Connectivity")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Device", "Compatible with")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("type", "Type")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Mic", "Microphone")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("noise", "Noise")));

            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArTitle", "Title Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArBran", "Brand Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDes", "Description Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arcolor", "Color Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArModel", "Model Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arconnectivity", "Connectivity Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("arDevice", "Compatible with Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("artype", "Type Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("arMic", "Microphone Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("arnoise", "Noise Arabic")));

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
                        setModel(i);
                        setConnectivity(i);
                         setCompatibleWith(i);
                        setType(i);
                        setMic(i);
                        setNoise(i);
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
            string type = "";
            if (Form1.OrganizedSheet[Micophone, row].Value.ToString() == "YES")
            {
                type = " Headset";
            }
            else
            {
                type = " Headphone";
            }
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase( Form1.OrganizedSheet[Brand, row].Value + " " + Form1.OrganizedSheet[Model, row].Value + " " +
                Form1.OrganizedSheet[Type, row].Value + type + ", " + Form1.OrganizedSheet[Colors, row].Value);
            Form1.BulkSheet[10, row].Value = "سماعة " + db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString()) + " من " +
                db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()) + " " + Form1.OrganizedSheet[Model, row].Value;
            if (CheckEnglish(Form1.BulkSheet[10, row].Value.ToString()))
            {
                Form1.BulkSheet[10, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setBrand(int row)
        {
            Form1.BulkSheet[1, row].Value = Form1.OrganizedSheet[Brand, row].Value;
            Form1.BulkSheet[11, row].Value = db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[11, row].Value.ToString()))
            {
                Form1.BulkSheet[11, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setDescription(int row)
        {
            string des = "", ArDes = "";
            try
            {

                string[] lines = Form1.OrganizedSheet[Extra, row].Value.ToString().Split('\n');
                foreach (string line in lines)
                {
                    des = des + "<p>" + line + "</p>";
                    ArDes = ArDes + "<p>" + db.getRecord(line) + "</p>";
                }
            }
            catch { }
            des = des + "<li>Brand: " + Form1.OrganizedSheet[Brand, row].Value+ "<li>Model: " + Form1.OrganizedSheet[Model, row].Value + "</li><li>Color: " + Form1.OrganizedSheet[Colors, row].Value + "</li><li>Type: " + Form1.OrganizedSheet[Type, row].Value +
"</li><li>Compatible with: " + Form1.OrganizedSheet[Device, row].Value + "</li>";
            ArDes = ArDes + "<li>العلامة التجارية: " +db.getRecord( Form1.OrganizedSheet[Brand, row].Value.ToString()) + "<li>الموديل: " +
                db.getRecord(Form1.OrganizedSheet[Model, row].Value.ToString())+ "<li>اللون: " + db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString()) + "</li><li>النوع: " +
                db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString()) + "</li><li>متوافق مع : " + db.getRecord(Form1.OrganizedSheet[Device, row].Value.ToString()) + "</li>";

            Form1.BulkSheet[2, row].Value = des;
            Form1.BulkSheet[12, row].Value = ArDes;
            if (CheckEnglish(ArDes))
            {
                Form1.BulkSheet[12, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setColor(int row)
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            Form1.BulkSheet[3, row].Value = textInfo.ToTitleCase(Form1.OrganizedSheet[Colors, row].Value.ToString());
            Form1.BulkSheet[13, row].Value = db.getRecord(Form1.OrganizedSheet[Colors, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[13, row].Value.ToString()))
            {
                Form1.BulkSheet[11, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            
        }
        private void setModel(int row)
        {
            Form1.BulkSheet[4, row].Value = Form1.OrganizedSheet[Model, row].Value;
            Form1.BulkSheet[14, row].Value = Form1.OrganizedSheet[Model, row].Value;
        }
        private void setConnectivity(int row)
        {
            Form1.BulkSheet[5, row].Value = Form1.OrganizedSheet[Connectivty, row].Value;
            Form1.BulkSheet[15, row].Value = db.getRecord(Form1.OrganizedSheet[Connectivty, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[15, row].Value.ToString()))
            {
                Form1.BulkSheet[15, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setCompatibleWith(int row)
        {
            Form1.BulkSheet[6, row].Value = Form1.OrganizedSheet[Device, row].Value;
            Form1.BulkSheet[16, row].Value = db.getRecord(Form1.OrganizedSheet[Device, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[16, row].Value.ToString()))
            {
                Form1.BulkSheet[16, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setType(int row)
        {
            Form1.BulkSheet[7, row].Value = Form1.OrganizedSheet[Type, row].Value;
            Form1.BulkSheet[17, row].Value = db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[17, row].Value.ToString()))
            {
                Form1.BulkSheet[17, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setMic(int row)
        {
            Form1.BulkSheet[8, row].Value = Form1.OrganizedSheet[Micophone, row].Value;
            Form1.BulkSheet[18, row].Value = db.getRecord(Form1.OrganizedSheet[Micophone, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[18, row].Value.ToString()))
            {
                Form1.BulkSheet[18, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setNoise(int row)
        {
            Form1.BulkSheet[9, row].Value = Form1.OrganizedSheet[Type, row].Value;
            Form1.BulkSheet[19, row].Value = db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString());
            if (CheckEnglish(Form1.BulkSheet[19, row].Value.ToString()))
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
