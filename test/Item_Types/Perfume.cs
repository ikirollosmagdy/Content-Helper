using System;
using System.Collections.Generic;
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



        public static int Brand = 1, Gender = 5, Size = 4, Type = 3, PerfumeName = 2,
            ExtraData = 6, FregFamily = 7;
        static bool isFirst = true;

        public static void createBulk()
        {



            if (isFirst)
            {
                //  Form1.BulkSheet.Columns.Add("Title", "Title");
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[0].HeaderText = "Title"));
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
                isFirst = false;
            }
            else
            {
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[0].HeaderText = "Title"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[1].HeaderText = "Brand"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[2].HeaderText = "Description"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[3].HeaderText = "Fregrance Type"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[4].HeaderText = "Size"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[5].HeaderText = "Target Group"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[6].HeaderText = "Fragrance Family"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[7].HeaderText = "Perfume Name"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[8].HeaderText = "Title AR"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[9].HeaderText = "Brand Ar"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[10].HeaderText = "Description(AR)"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[11].HeaderText = "Fregrance Type AR"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[12].HeaderText = "Size AR"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[13].HeaderText = "Target Group AR"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[14].HeaderText = "Fragrance Family(AR)"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[15].HeaderText = "Perfume Name(AR)"));
            }

            for (int i = 0; i < Form1.OrganizedSheet.RowCount; i++)
            {
                setTitle(i);
                setBrand(i);
                setDescription(i);
                setType(i);
                setSize(i);
                setGender(i);
                setFragFamiley(i);
                setPerfumeName(i);


            }


        }


        public static void getGender()
        {
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.RowCount = Form1.OrganizedSheet.RowCount));
            int row, col;
           
            //   Form1.OrganizedSheet.RowCount = Form1.Sheet.RowCount;
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[0].HeaderText = "Row Number"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[Brand].HeaderText = "Brand"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[PerfumeName].HeaderText = "Perfume Name"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[Gender].HeaderText = "Gender"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[Size].HeaderText = "Size"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[Type].HeaderText = "Fregrance Type"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[ExtraData].HeaderText = "Description"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[FregFamily].HeaderText = "Fregrance Family"));
            for (row = 0; row < Form1.Sheet.RowCount; row++)
            {
                for (col = 0; col < Form1.Sheet.ColumnCount; col++)
                {
                    try
                    {
                        Regex regexSize = new Regex(@"(\d{2,3} ?[Mm][lL]?)");
                        Match matchSize = regexSize.Match(Form1.Sheet.Rows[row].Cells[col].Value.ToString());



                        if (matchSize.Success)
                        {

                            Form1.OrganizedSheet.Rows[row].Cells[Size].Value = matchSize.Value.Replace(" ", "").ToLower();
                          
                        }
                    }
                    catch (Exception) { }
                    try
                    {

                        Regex regexGender = new Regex(@"(([Ww]o)?[Mm][ae]n)|([Uu]ni(sex)?)|([Bb]oys?)|([Gg]irls?)");
                        MatchCollection matchGender = regexGender.Matches(Form1.Sheet.Rows[row].Cells[col].Value.ToString());
                        if (matchGender.Count > 1)
                        {
                            Form1.OrganizedSheet.Rows[row].Cells[Gender].Value = getReplacement(matchGender[0].Value) + ", " + getReplacement(matchGender[1].Value);
                           
                        }
                        else
                        {
                            //  Form1.OrganizedSheet.Rows[row].Cells[Gender].Value = textInfo.ToTitleCase(matchGender[0].Value);
                            Form1.OrganizedSheet.Rows[row].Cells[Gender].Value = getReplacement(matchGender[0].Value);
                          
                        }
                        
                    }
                    catch (Exception) { }
                    try
                    {

                        Regex regType = new Regex(@"((?=[Ee]au)( ?\w+ ?\w+ ?\w{5,8}))|((?=[Pp]erf)(\w{6} ?\w+ ?\w{3,6}))|(([A-Z]\w+ ?\w+ ?)(?=[Pp]arf)\w{6})|([Oo]ud ?)|([Ee][Dd][TtPpCc] ?)");
                        Match matchType = regType.Match(Form1.Sheet.Rows[row].Cells[col].Value.ToString());
                        if (matchType.Success)
                        {
                            
                            addTypeAttrib(row, Type, getReplacement(matchType.Value));

                           
                        }





                    }
                    catch (Exception)
                    {



                    }
                    try
                    {
                        
                        Form1.OrganizedSheet.Rows[row].Cells[0].Value = Convert.ToString(row + 2);
                    
                    }
                    catch (Exception)
                    {

                       
                    }

                   


                    addAttrib(row);
                    if (Form1.OrganizedSheet[Type, row] == null || Form1.OrganizedSheet[Type, row].Value == null || Form1.OrganizedSheet[Type, row].Value.ToString() == "")
                    {
                        addTypeAttrib(row, Type, null);

                    }
                }
            }

        }




        public static void setTitle(int row)
        {


            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            Database db = new Database();

            string title = Form1.OrganizedSheet.Rows[row].Cells[PerfumeName].Value + " By " + Form1.OrganizedSheet.Rows[row].Cells[Brand].Value
              + " For " + Form1.OrganizedSheet.Rows[row].Cells[Gender].Value + " - " + Form1.OrganizedSheet.Rows[row].Cells[Type].Value
               + ", " + Form1.OrganizedSheet.Rows[row].Cells[Size].Value;
            Form1.BulkSheet.Rows[row].Cells[0].Value = textInfo.ToTitleCase(title);
            Form1.BulkSheet[8, row].Value = "عطر " + db.getRecord(Form1.OrganizedSheet[PerfumeName, row].Value.ToString()) + " ل"
                + db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString()) + " من " + db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()) +
                " - " + db.getRecord(Form1.OrganizedSheet[Type, row].Value.ToString()) + "، " + db.getRecord(Form1.OrganizedSheet[Size, row].Value.ToString());

        }

        public static void setDescription(int row)
        {
            Database db = new Database();
            string Des = "<p>" + Form1.OrganizedSheet.Rows[row].Cells[ExtraData].Value
                + "</p> <p><strong>Product Features:</strong></p> <ul> <li>Brand: "
                + Form1.OrganizedSheet.Rows[row].Cells[Brand].Value + "</li> <li>Target Group: "
                + Form1.OrganizedSheet.Rows[row].Cells[Gender].Value + "</li> <li>Fragrance Type: "
                + Form1.OrganizedSheet.Rows[row].Cells[Type].Value + "</li> <li>Perfume Name: "
                + Form1.OrganizedSheet.Rows[row].Cells[PerfumeName].Value + "</li> <li>Size: "
                + Form1.OrganizedSheet.Rows[row].Cells[Size].Value + "</li> </ul>";
            Form1.BulkSheet.Rows[row].Cells[2].Value = Des;
            Form1.BulkSheet.Rows[row].Cells[10].Value = "<p>" + db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[ExtraData].Value.ToString())
                + "</p> <p><strong>خصائص المنتج:</strong></p> <ul> <li>العلامة التجارية: "
                + db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[Brand].Value.ToString()) + "</li> <li>النوع: "
                + db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[Gender].Value.ToString()) + "</li> <li>النوع العطر: "
                + db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[Type].Value.ToString()) + "</li> <li>اسم العطر: "
                + db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[PerfumeName].Value.ToString()) + "</li> <li>الحجم: "
                + db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[Size].Value.ToString()) + "</li> </ul>";


        }
        public static void setBrand(int row)
        {
            Database db = new Database();
            Form1.BulkSheet.Rows[row].Cells[1].Value = Form1.OrganizedSheet.Rows[row].Cells[Brand].Value;
            Form1.BulkSheet.Rows[row].Cells[9].Value = db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[Brand].Value.ToString());
        }
        public static void setType(int row)
        {
            Form1.BulkSheet.Rows[row].Cells[3].Value = Form1.OrganizedSheet.Rows[row].Cells[Type].Value;

            if (Form1.OrganizedSheet.Rows[row].Cells[Type].Value != null)
            {
                Database db = new Database();

                Form1.BulkSheet.Rows[row].Cells[11].Value = db.getRecord(Form1.BulkSheet.Rows[row].Cells[3].Value.ToString());




            }
        }
        public static void setSize(int row)
        {
            Form1.BulkSheet.Rows[row].Cells[4].Value = Form1.OrganizedSheet.Rows[row].Cells[Size].Value;
            if (Form1.OrganizedSheet.Rows[row].Cells[Size].Value != null)
            {
                Database db = new Database();
                Form1.BulkSheet.Rows[row].Cells[12].Value = db.getRecord(Form1.BulkSheet.Rows[row].Cells[4].Value.ToString());
            }

        }
        public static void setGender(int row)
        {
            Form1.BulkSheet.Rows[row].Cells[5].Value = Form1.OrganizedSheet.Rows[row].Cells[Gender].Value;

            if (Form1.OrganizedSheet.Rows[row].Cells[Gender].Value != null)
            {
                Database db = new Database();
                Form1.BulkSheet.Rows[row].Cells[13].Value = db.getRecord(Form1.BulkSheet.Rows[row].Cells[5].Value.ToString());
            }

        }

        public static void setFragFamiley(int row)
        {
            Database db = new Database();
            Form1.BulkSheet.Rows[row].Cells[6].Value = Form1.OrganizedSheet.Rows[row].Cells[FregFamily].Value;
            Form1.BulkSheet[14, row].Value = db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[FregFamily].Value.ToString());
        }
        public static void setPerfumeName(int row)
        {
            Database db = new Database();
            Form1.BulkSheet.Rows[row].Cells[7].Value = Form1.OrganizedSheet.Rows[row].Cells[PerfumeName].Value;
            Form1.BulkSheet[15, row].Value = db.getRecord(Form1.OrganizedSheet.Rows[row].Cells[PerfumeName].Value.ToString());
        }



        public static string getReplacement(String Text)
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

        public static void addAttrib(int rows)
        {
                       
            string[] datasource = { "Floral & Fruity", "Fresh & Zesty", "Oriental & Spicy", "Oriental Fruity", "Woody & Musky", "Woody & Spicy" };
            DataGridViewComboBoxCell combo = new DataGridViewComboBoxCell();
            combo.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            combo.FlatStyle = FlatStyle.Popup;
            combo.DataSource = datasource.ToList();
            Form1.OrganizedSheet[7, rows] = combo;
           

        }
        public static void addTypeAttrib(int row,int col,string data)
        {
            List<string> list = new List<string>();
          
            DataGridViewComboBoxCell combo = new DataGridViewComboBoxCell();
            combo.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            combo.FlatStyle = FlatStyle.Popup;
            addItems(list);
            int ID = list.FindIndex(xBas => xBas.Equals(data, StringComparison.OrdinalIgnoreCase));
            if (ID != -1) {
                combo.DataSource = list;
                Form1.OrganizedSheet[col, row] = combo;
                Form1.OrganizedSheet[col, row].Value = list[ID];
            }
            else {



                if (data == null)
                {
                    combo.DataSource = list;
                    Form1.OrganizedSheet[col, row] = combo;
                }
                else
                {
                    list.Add(data);
                    combo.DataSource = list;
                    Form1.OrganizedSheet[col, row] = combo;

                    Form1.OrganizedSheet[col, row].Value = data;
                }
            }
        }

        public static void dropdown(DataGridViewEditingControlShowingEventArgs e)
        {
            if (Form1.OrganizedSheet.CurrentCell.OwningColumn.HeaderText != "Fregrance Type")
            {
                return;
            }
            else { 
                TextBox autoText = e.Control as TextBox;
                if (autoText != null)
                {
                    autoText.AutoCompleteMode = AutoCompleteMode.Suggest;
                    autoText.AutoCompleteSource = AutoCompleteSource.CustomSource;
                    AutoCompleteStringCollection DataCollection = new AutoCompleteStringCollection();
                  //  addItems(DataCollection);
                    autoText.AutoCompleteCustomSource = DataCollection;

                }
               


                
            }
        }
              public static void addItems(List<string> col)
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
