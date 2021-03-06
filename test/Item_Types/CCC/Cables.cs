﻿using System;
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
    public class Cables
    {
        Database db = new Database();
        Common_Use common = new Common_Use();
        int Brand = 1, Model = 2, Type = 3, Colors = 4, Length = 5, Device = 6,
            Link = 7, Price = 8, Quantity = 9, UnTranslatedCount = 0;

        private void setupTable()
        {
            DataGridViewComboBoxColumn typeColumn = new DataGridViewComboBoxColumn();
            typeColumn.HeaderText = "Type";
            typeColumn.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            typeColumn.FlatStyle = FlatStyle.Popup;
            typeColumn.Items.AddRange("Adapter", "Audio Cable", "Audio Splitter", "Cable", "Power cord", "VGA Splitter");

            DataGridViewComboBoxColumn Device = new DataGridViewComboBoxColumn();
            Device.HeaderText = "Compatible with";
            Device.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            Device.FlatStyle = FlatStyle.Popup;
            Device.Items.AddRange("Camcorders", "Computers - PCs", "Digital Cameras", "Game Consoles", "Laptops & Notebooks", "Mobile Phones", "MP3 Players & MP4 Players", "Multi", "Printers", "Projectors", "Routers", "Smart Watches", "Speakers", "TV's");

            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("NO", "No.")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Model", "Model")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(typeColumn)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Color", "Color")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Length", "Length")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(Device)));
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
                for (col = 1; col < Form1.Sheet.ColumnCount; col++)
                {
                    try
                    {
                        Form1.OrganizedSheet[0, row].Value = Convert.ToString(row + 1);

                    }
                    catch { }

                    if (Form1.Sheet.Columns[col].HeaderText.ToLower() == "brand")
                    {
                        Form1.OrganizedSheet[Brand, row].Value = Form1.Sheet[col, row].Value.ToString().Trim();
                    }

                    try
                    {
                        if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("length")) {
                            Form1.OrganizedSheet[Length, row].Value = Form1.Sheet[col, row].Value.ToString().Replace(" ", "").ToLower();
                        }
                        else
                        {
                            Regex regexSize = new Regex(@"(\d{1,2} ?[Mm])|(\d{2,3} ?[Cc][Mm])");
                            Match matchSize = regexSize.Match(Form1.Sheet.Rows[row].Cells[col].Value.ToString());



                            if (matchSize.Success)
                            {

                                Form1.OrganizedSheet.Rows[row].Cells[Length].Value = matchSize.Value.Replace(" ", "").ToLower();

                            }
                        }
                    }
                    catch { }
                    try
                    {
                        if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("color")) {

                            Form1.OrganizedSheet[Colors, row].Value = Form1.Sheet[col, row].Value.ToString().Trim();
                        }
                        else
                        {
                            Regex regexColor = new Regex(@"([Bb]eige|[Bb]lack|[Bb]lue|[Bb]rown|[Cc]lear|[Gg]old|[Gg]reen|[Gg]rey|[Mm]ulti ?[Cc]olor|[Oo]ffWhite|[Oo]range|[Pp]ink|[Pp]urple|[Rr]ed|[Ss]ilver|[Tt]urquoise|[Ww]hite|[Yy]ellow)");
                            MatchCollection matchColor = regexColor.Matches(Form1.Sheet[col, row].Value.ToString());


                            if (matchColor.Count > 0)
                            {
                                Form1.OrganizedSheet[Colors, row].Value = matchColor[0].Value.Trim();


                            }
                        }
                    }
                    catch { }
                    try
                    {
                        for (int y = 1; y < Form1.Sheet.ColumnCount; y++)
                        {
                            if (Form1.Sheet.Columns[y].HeaderText.ToLower().Contains("model") || Form1.Sheet.Columns[y].HeaderText.ToLower().Contains("code")) {

                                Form1.OrganizedSheet[Model, row].Value = Form1.Sheet[y, row].Value.ToString().Trim();
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
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Device", "Device")));

            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArTitle", "Title Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArBran", "Brand Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDes", "Description Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ARType", "Type Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Ar Device", "Device Arabic")));

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
                        setCompatibleWith(i);

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
                ArTitle = " من " + db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()) + " " + Form1.OrganizedSheet[Model, row].Value + " ";
            }
            title = title + Form1.OrganizedSheet[Type, row].Value + ", " + Form1.OrganizedSheet[Length, row].Value + ", " + Form1.OrganizedSheet[Colors, row].Value;
            ArTitle = db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString()) + ArTitle + " ،" + db.getRecord(Form1.OrganizedSheet[Length, row].Value.ToString()) + " ،" +
                common.getColorMulti(Form1.OrganizedSheet[Colors, row].Value.ToString());
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase(title);
            Form1.BulkSheet[5, row].Value = ArTitle;
            if (common.CheckEnglish(Form1.BulkSheet[5, row].Value.ToString()))
            {
                Form1.BulkSheet[5, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }

        }
        private void setBrand(int row)
        {
            Form1.BulkSheet[1, row].Value = Form1.OrganizedSheet[Brand, row].Value;
            Form1.BulkSheet[6, row].Value = db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[6, row].Value.ToString()))
            {
                Form1.BulkSheet[6, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setDescription(int row)
        {
            string brand = "",arBrand="";
            if (Form1.OrganizedSheet[Brand, row].Value.ToString().ToLower().Trim() != "other")
            {
                brand = "<ul> <li>Brand :" + Form1.OrganizedSheet[Brand, row].Value+ "</li>";
                arBrand= "<ul> <li>العلامة التجارية :" + db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()) + "</li>"; 
            }
            else {
                brand = "<ul>";
                arBrand = "<ul>";
            }

            Form1.BulkSheet[2, row].Value = brand + "<li>Color :" +
                Form1.OrganizedSheet[Colors, row].Value + "</li> <li>Length :" + Form1.OrganizedSheet[Length, row].Value +
                "</li> <li>Compatible with :" + Form1.OrganizedSheet[Device, row].Value + "</li> </ul>";
            Form1.BulkSheet[7, row].Value = arBrand+"<li>اللون :" +
               common.getColorMulti(Form1.OrganizedSheet[Colors, row].Value.ToString()) + "</li> <li>الطول :" + db.getRecord(Form1.OrganizedSheet[Length, row].Value.ToString()) +
                "</li> <li>متوافق مع :" + db.getRecord(Form1.OrganizedSheet[Device, row].Value.ToString()) + "</li> </ul>";
            if (common.CheckEnglish(Form1.BulkSheet[7, row].Value.ToString()))
            {
                Form1.BulkSheet[7, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }

        }
        private void setType(int row)
        {
            if (Form1.OrganizedSheet[Type, row].Value.ToString() == "Cable" || Form1.OrganizedSheet[Type, row].Value.ToString() == "Adapter")
            {
                Form1.BulkSheet[3, row].Value = Form1.OrganizedSheet[Type, row].Value + "s";
            }
            else
            {
                Form1.BulkSheet[3, row].Value = Form1.OrganizedSheet[Type, row].Value;
            }
            Form1.BulkSheet[8, row].Value = db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[8, row].Value.ToString()))
            {
                Form1.BulkSheet[8, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setCompatibleWith(int row)
        {
            Form1.BulkSheet[4, row].Value = Form1.OrganizedSheet[Device, row].Value;
            Form1.BulkSheet[9, row].Value = db.getRecord(Form1.OrganizedSheet[Device, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[9, row].Value.ToString()))
            {
                Form1.BulkSheet[9, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setLink(int row)
        {
            Form1.BulkSheet[10, row].Value = Form1.OrganizedSheet[Link, row].Value;
        }
        private void setPrice(int row)
        {
            Form1.BulkSheet[11, row].Value = Form1.OrganizedSheet[Price, row].Value;
        }
        private void setQuantity(int row)
        {
            Form1.BulkSheet[12, row].Value = Form1.OrganizedSheet[Quantity, row].Value;
        }

       
    }
}
