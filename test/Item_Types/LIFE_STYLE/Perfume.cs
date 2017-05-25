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
    public class Perfume

    {

        Database db = new Database();
        Common_Use common = new Common_Use();
        int Brand = 1, Gender = 5, Size = 4, Type = 3, PerfumeName = 2,
            ExtraData = 7, FregFamily = 6,UnTranslatedCount=0,Link=8,Price=9,Quantity=10;
       

        private  void setupTable()

        {
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));

            DataGridViewComboBoxColumn com = new DataGridViewComboBoxColumn();
            com.HeaderText = "Fregrance Type";
            com.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            com.FlatStyle = FlatStyle.Popup;
            com.Items.AddRange("Eau de Cologne", "Eau de Parfum", "Eau de Splash", "Eau de Toilette", "Esprit de Parfum", "Extrait De Parfum"
                , "Oud", "Perfume Mist", "Perfume Oil");
            DataGridViewComboBoxColumn Freg = new DataGridViewComboBoxColumn();
            Freg.HeaderText = "Fregrance Familey";
            Freg.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            Freg.FlatStyle = FlatStyle.Popup;
            Freg.Items.AddRange("Floral & Fruity", "Fresh & Zesty", "Oriental & Spicy", "Oriental Fruity", "Woody & Musky", "Woody & Spicy");

            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("NO", "No.")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("PerfumeName", "Perfume Name")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(com)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Size", "Size")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Gender", "Gender")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(Freg)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("De", "Description")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Link", "Image Url")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Price", "Price")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Qunatity", "Quantity")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.RowCount = Form1.Sheet.RowCount));
            foreach(DataGridViewColumn co in Form1.OrganizedSheet.Columns)
            {
                co.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        public  void createBulk()
        {
           
                UnTranslatedCount = 0;
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Clear()));
                Form1.txtUntranslated.GetCurrentParent().Invoke(new Action(() => Form1.txtUntranslated.Text = UnTranslatedCount.ToString()));

                //  Form1.BulkSheet.Columns.Add("Title", "Title");
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Title", "Title")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Bran", "Brand")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Des", "Description")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Type", "Fregrance Type")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Size", "Size")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Gender", "Target Group")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Familey", "Fragrance Family")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("FragName", "Perfume Name")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("arTitle", "Arabic title")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("arBrand", "Brand Ar")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("arDes", "Description(AR)")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("arType", "Fregrance Type AR")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Sizear", "Size Ar")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("GenderAr", "Target Group AR")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("FmileyAr", "Fragrance Family(AR)")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("FragNameAr", "Perfume Name(AR) ")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Link", "Link")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Price", "Price")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Qunt", "Quntity")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ean", "Suggested EAN")));

            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.RowCount = Form1.OrganizedSheet.RowCount));
            

            for (int i = 0; i < Form1.OrganizedSheet.RowCount - 1; i++)
                {
                try
                {
                    if (Form1.OrganizedSheet.Rows[i].DefaultCellStyle.BackColor != Color.HotPink)
                    {
                        setTitle(i);
                        setBrand(i);
                        setDescription(i);
                        setType(i);
                        setSize(i);
                        setGender(i);
                        setFragFamiley(i);
                        setPerfumeName(i);
                        setLink(i);
                        setPrice(i);
                        setQuantity(i);
                        Form1.txtUntranslated.GetCurrentParent().Invoke(new Action(() => Form1.txtUntranslated.Text = UnTranslatedCount.ToString()));
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error: " + ex.Message);
                }
                }

           
        }


        public  void getGender()
        {
            setupTable();
            int row, col;
            
            for (row = 0; row < Form1.Sheet.RowCount; row++)
            {
                for (col = 0; col < Form1.Sheet.ColumnCount; col++)
                {
                    if (Form1.Sheet.Columns[col].HeaderText.ToLower() == "brand")
                    {
                        Form1.OrganizedSheet[Brand, row].Value = Form1.Sheet[col, row].Value;
                    }
                    try
                    {
                        if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("size"))
                        {
                            Form1.OrganizedSheet[Size, row].Value = Form1.Sheet[col, row].Value;
                        }
                        else
                        {
                            Regex regexSize = new Regex(@"(\d{2,3} ?[Mm][lL]?)");
                            Match matchSize = regexSize.Match(Form1.Sheet.Rows[row].Cells[col].Value.ToString());



                            if (matchSize.Success)
                            {

                                Form1.OrganizedSheet.Rows[row].Cells[Size].Value = matchSize.Value.Replace(" ", "").ToLower();

                            }
                        }
                    }
                    catch { }

                    if (Form1.Sheet.Columns[col].HeaderText.ToLower().Contains("name"))
                    {
                        Form1.OrganizedSheet[PerfumeName, row].Value = Form1.Sheet[col, row].Value;
                    }

                    try
                    {

                        Regex regexGender = new Regex(@"(([Ww]o)?[Mm][ae]n)|([Uu]ni(sex)?)|([Bb]oys?)|([Gg]irls?)");
                        MatchCollection matchGender = regexGender.Matches(Form1.Sheet.Rows[row].Cells[col].Value.ToString());
                        if (matchGender.Count > 1)
                        {
                            Form1.OrganizedSheet.Rows[row].Cells[Gender].Value = common.getReplacement(matchGender[0].Value) + ", " + common.getReplacement(matchGender[1].Value);

                        }
                        else
                        {
                            //  Form1.OrganizedSheet.Rows[row].Cells[Gender].Value = textInfo.ToTitleCase(matchGender[0].Value);
                            Form1.OrganizedSheet.Rows[row].Cells[Gender].Value = common.getReplacement(matchGender[0].Value);

                        }

                    }
                    catch (Exception) { }
                    try
                    {

                        Regex regType = new Regex(@"((?=[Ee]au)( ?\w+ ?\w+ ?\w{5,8}))|((?=[Pp]erf)(\w{6} ?\w+ ?\w{3,6}))|(([A-Z]\w+ ?\w+ ?)(?=[Pp]arf)\w{6})|([Oo]ud ?)|([Ee][Dd][TtPpCc] ?)");
                        Match matchType = regType.Match(Form1.Sheet.Rows[row].Cells[col].Value.ToString());
                        if (matchType.Success)
                        {

                            addTypeAttrib(row, Type, common.getReplacement(matchType.Value));


                        }





                    }
                    catch (Exception)
                    {



                    }
                    try
                    {

                        Form1.OrganizedSheet.Rows[row].Cells[0].Value = Convert.ToString(row + 1);

                    }
                    catch (Exception) { }
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




                    if (Form1.OrganizedSheet[Type, row] == null || Form1.OrganizedSheet[Type, row].Value == null || Form1.OrganizedSheet[Type, row].Value.ToString() == "")
                    {
                        addTypeAttrib(row, Type, null);

                    }

                }
            }

        }




          void setTitle(int row)
        {


            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            Database db = new Database();

            string title = Form1.OrganizedSheet.Rows[row].Cells[PerfumeName].Value + " By " + Form1.OrganizedSheet.Rows[row].Cells[Brand].Value
              + " For " + Form1.OrganizedSheet.Rows[row].Cells[Gender].Value + " - " + Form1.OrganizedSheet.Rows[row].Cells[Type].EditedFormattedValue
               + ", " + Form1.OrganizedSheet.Rows[row].Cells[Size].Value;
                 Form1.BulkSheet.Rows[row].Cells[0].Value = textInfo.ToTitleCase(title);

            Form1.BulkSheet[8, row].Value = "عطر " + db.getRecord(Form1.OrganizedSheet[PerfumeName, row].Value.ToString()) + " ل"
                + db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString()) + " من " + db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()) +
                " - " + db.getRecord(Form1.OrganizedSheet[Type, row].EditedFormattedValue.ToString()) + "، " + Form1.OrganizedSheet[Size, row].Value.ToString().Replace("ml","مل");

            if(common.CheckEnglish(Form1.BulkSheet[8, row].Value.ToString()))
            {
                Form1.BulkSheet[8, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[8, row].Style.BackColor = Color.White;
            }

            //19
            Form1.BulkSheet[19, row].Value = db.getEAN(Form1.BulkSheet[0, row].Value.ToString());
        }

          void setDescription(int row)
        {
            Database db = new Database();
            string Des = "<p>" + Form1.OrganizedSheet.Rows[row].Cells[ExtraData].Value
                + "</p> <p><b>Product Features:</b></p> <ul> <li>Brand: "
                + Form1.OrganizedSheet.Rows[row].Cells[Brand].Value + "</li> <li>Target Group: "
                + Form1.OrganizedSheet.Rows[row].Cells[Gender].Value + "</li> <li>Fragrance Type: "
                + Form1.OrganizedSheet.Rows[row].Cells[Type].EditedFormattedValue + "</li> <li>Perfume Name: "
                + Form1.OrganizedSheet.Rows[row].Cells[PerfumeName].Value + "</li> <li>Size: "
                + Form1.OrganizedSheet.Rows[row].Cells[Size].Value + "</li> </ul>";
            Form1.BulkSheet.Rows[row].Cells[2].Value = Des;
            Form1.BulkSheet.Rows[row].Cells[10].Value = "<p>" + db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[ExtraData].Value.ToString())
                + "</p> <p><b>خصائص المنتج:</b></p> <ul> <li>العلامة التجارية: "
                + db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[Brand].Value.ToString()) + "</li> <li>النوع: "
                + db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[Gender].Value.ToString()) + "</li> <li>النوع العطر: "
                + db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[Type].EditedFormattedValue.ToString()) + "</li> <li>اسم العطر: "
                + db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[PerfumeName].Value.ToString()) + "</li> <li>الحجم: "
                + Form1.OrganizedSheet.Rows[row].Cells[Size].Value.ToString().Replace("ml","مل") + "</li> </ul>";
            if (common.CheckEnglish(Form1.BulkSheet.Rows[row].Cells[10].Value.ToString()))
            {
                Form1.BulkSheet[10, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[10, row].Style.BackColor = Color.White;
            }

        }
          void setBrand(int row)
        {
            Database db = new Database();
            Form1.BulkSheet.Rows[row].Cells[1].Value = Form1.OrganizedSheet.Rows[row].Cells[Brand].Value;
            Form1.BulkSheet.Rows[row].Cells[9].Value = db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[Brand].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[9, row].Value.ToString()))
            {
                Form1.BulkSheet[9, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[9, row].Style.BackColor = Color.White;
            }
        }
          void setType(int row)
        {
            Form1.BulkSheet.Rows[row].Cells[3].Value = Form1.OrganizedSheet.Rows[row].Cells[Type].Value;

            if (Form1.OrganizedSheet.Rows[row].Cells[Type].Value != null)
            {
                Database db = new Database();

                Form1.BulkSheet.Rows[row].Cells[11].Value = db.getRecord(Form1.BulkSheet.Rows[row].Cells[3].Value.ToString());

                if (common.CheckEnglish(Form1.BulkSheet[11, row].Value.ToString()))
                {
                    Form1.BulkSheet[11, row].Style.BackColor = Color.Yellow;
                    UnTranslatedCount++;
                }
                else
                {
                    Form1.BulkSheet[11, row].Style.BackColor = Color.White;
                }

            }
        }
          void setSize(int row)
        {
            Form1.BulkSheet.Rows[row].Cells[4].Value = Form1.OrganizedSheet.Rows[row].Cells[Size].Value.ToString().Replace(" ","");
            if (Form1.OrganizedSheet.Rows[row].Cells[Size].Value != null)
            {
                
                Form1.BulkSheet.Rows[row].Cells[12].Value = Form1.BulkSheet.Rows[row].Cells[4].Value.ToString().Replace("ml","مل");
                
            }

        }
          void setGender(int row)
        {
            Form1.BulkSheet.Rows[row].Cells[5].Value = Form1.OrganizedSheet.Rows[row].Cells[Gender].Value;

            if (Form1.OrganizedSheet.Rows[row].Cells[Gender].Value != null)
            {
                Database db = new Database();
                Form1.BulkSheet.Rows[row].Cells[13].Value = db.getRecord(Form1.BulkSheet.Rows[row].Cells[5].Value.ToString());
                if (common.CheckEnglish(Form1.BulkSheet[13, row].Value.ToString()))
                {
                    Form1.BulkSheet[13, row].Style.BackColor = Color.Yellow;
                    UnTranslatedCount++;
                }
                else
                {
                    Form1.BulkSheet[13, row].Style.BackColor = Color.White;
                }
            }

        }

          void setFragFamiley(int row)
        {
            Database db = new Database();
            Form1.BulkSheet.Rows[row].Cells[6].Value = Form1.OrganizedSheet.Rows[row].Cells[FregFamily].EditedFormattedValue;
            Form1.BulkSheet[14, row].Value = db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[FregFamily].EditedFormattedValue.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[14, row].Value.ToString()))
            {
                Form1.BulkSheet[14, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[14, row].Style.BackColor = Color.White;
            }
        }
          void setPerfumeName(int row)
        {
            Database db = new Database();
            Form1.BulkSheet.Rows[row].Cells[7].Value = Form1.OrganizedSheet.Rows[row].Cells[PerfumeName].Value;
            Form1.BulkSheet[15, row].Value = db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[PerfumeName].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[15, row].Value.ToString()))
            {
                Form1.BulkSheet[15, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
            else
            {
                Form1.BulkSheet[15, row].Style.BackColor = Color.White;
            }
        }

          void setLink(int row)
        {
            Form1.BulkSheet[16, row].Value = Form1.OrganizedSheet[Link, row].Value;
        }
          void setPrice(int row)
        {
            Form1.BulkSheet[17, row].Value = Form1.OrganizedSheet[Price, row].Value;
        }
          void setQuantity(int row)
        {
            Form1.BulkSheet[18, row].Value = Form1.OrganizedSheet[Quantity, row].Value;
        }


        

        
        
          void addTypeAttrib(int row, int col, string data)
        {
            List<string> list = new List<string>();

           
            addItems(list);
            int ID = list.FindIndex(xBas => xBas.Equals(data, StringComparison.OrdinalIgnoreCase));
            if (ID != -1)
            {
               
                Form1.OrganizedSheet[col, row].Value = list[ID];
            }
           
        }

      
          void addItems(List<string> col)
        {
            col.Add("Eau de Cologne");
            col.Add("Eau de Parfum");
            col.Add("Eau de Splash");
            col.Add("Eau de Toilette");
            col.Add("Esprit de Parfum");
            col.Add("Extrait De Parfum");
            col.Add("Oud");
            col.Add("Perfume Mist");
            col.Add("Perfume Oil");
        }

      
        
    }

    

}
