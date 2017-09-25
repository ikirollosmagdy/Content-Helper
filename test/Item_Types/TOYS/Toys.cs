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
    class Toys
    {
        Database db = new Database();
        TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
        Common_Use common = new Common_Use();
        int Brand = 1, Toy_name = 2,Desc = 3, category = 4, Gender = 5, Age = 6,
           Link = 7, Price = 8, Quantity = 9, UnTranslatedCount = 0;
        private void setupTable()
        {

            DataGridViewComboBoxColumn Category_Column = new DataGridViewComboBoxColumn();
            Category_Column.HeaderText = "Toy Category";
            Category_Column.DisplayStyle = DataGridViewComboBoxDisplayStyle.ComboBox;
            Category_Column.FlatStyle = FlatStyle.Popup;
            Category_Column.Items.AddRange("Action Figures", "Activity & Amusement", "Animal Kingdom", "Arts & Crafts", "Balls", "Bath Toys", "Battery Operated & Wind-Up", "Beach Toys", "Blaster Toys", "Board & Card Games", "Boats", "Brain Teasers", "Bubble Guns", "Bubbles Bottels", "Buzzers for Contests and Games", "Car Racetracks", "Cars", "Collection Cards", "Construction, Building Sets & Blocks", "Counting Frames", "Decoration Figures", "Dice & Dice Games", "Die Casts", "Dolls", "Dominoes", "Educational Computers", "Educational Toys", "Electronic Toys", "Face Paints", "Guns", "Hammering & Nailing Toys", "Helicopters", "Hobby, Models & Trains", "Jigsaws & Puzzles", "Kids Art Furniture", "Kites & Flight Toys", "Light, Sound & Music Toys", "Magic Tricks", "Makeup Toys", "Marbles", "Masks", "Moneybox", "Motor Toys", "Mystery Games", "Paper Baloons", "Planes", "Play Tents", "Pretend & Dress Up", "Puppet & Puppet Theatre", "Remote Controlled Toys", "Robotics", "Sandpit & Sand Toys", "Scaled Models", "Smart Toys", "Stress Relief Toys", "Stuffed Toys", "Temporary Tattoos", "Toy Target Games", "Toys Accessories", "Vehicles", "Walkie Talkies", "Water Guns", "Yoyo");
            

            


            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("NO", "No.")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Name", "Toy Name")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Desc", "Description")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add(Category_Column)));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Gender", "Gender")));

            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("age", "Age number")));
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
                            if (Form1.Sheet.Columns[y].HeaderText.ToLower().Contains("name") )
                            {
                                Form1.OrganizedSheet[Toy_name, row].Value = Form1.Sheet[y, row].Value;
                                break;
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
                            if (matchGender.Value.ToLower() == "woman")
                            {
                                Form1.OrganizedSheet.Rows[row].Cells[Gender].Value = "Women";
                            }
                            else
                            {
                                Form1.OrganizedSheet.Rows[row].Cells[Gender].Value = common.getReplacement(matchGender.Value);
                            }
                        }


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
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("category", "Toy Category")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("gender", "Gender")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Age", "Age")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("recom", "Recommended age")));
           
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArTitle", "Title Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArBran", "Brand Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArDes", "Description Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arcategory", "Toy Category Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Argender", "Gender Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("ArAge", "Age Arabic")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Arrecom", "Recommended age Arabic")));

            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Link", "Link")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Price", "Price")));
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Quan", "Quantity")));
            

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
                        setCategory(i);
                        setGender(i);
                        setAge(i);
                        setRecommendedAge(i);
                        

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
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase(string.Format("{0} {1} for {2}",
               Form1.OrganizedSheet[Brand, row].Value, Form1.OrganizedSheet[Toy_name, row].Value, Form1.OrganizedSheet[Gender, row].Value));
            Form1.BulkSheet[7, row].Value = string.Format("{0} من {1} لل{2}", db.getRecord(Form1.OrganizedSheet[Toy_name, row].Value.ToString()),
                db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString()));
            if (common.CheckEnglish(Form1.BulkSheet[7, row].Value.ToString()))
            {
                Form1.BulkSheet[7, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setBrand(int row)
        {
            Form1.BulkSheet[1, row].Value = Form1.OrganizedSheet[Brand, row].Value;
            Form1.BulkSheet[8, row].Value = db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[8, row].Value.ToString()))
            {
                Form1.BulkSheet[8, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setDescription(int row)
        {
            Form1.BulkSheet[2, row].Value = string.Format("<p>{0}</p><p><b>Product Features:</b></p><ul><li>Brand:{1}</li><li>Toy Category:{2}</li><li>Recommended Age:{3}</li><li>Targeted Group:{4}</li></ul>",
                Form1.OrganizedSheet[Desc, row].Value, Form1.OrganizedSheet[Brand, row].Value, Form1.OrganizedSheet[category, row].Value, Form1.OrganizedSheet[Age, row].Value, Form1.OrganizedSheet[Gender, row].Value);
            Form1.BulkSheet[9, row].Value = string.Format("<p>{0}</p><p><b>خصائص المنتج:</b></p><ul><li>العلامة  التجارية:{1}</li><li>نوع اللعبة:{2}</li><li>الفئة العمرية:3}</li><li>المجموعة المستهدفة:{4}</li></ul>",
               db.getRecord(Form1.OrganizedSheet[Desc, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[Brand, row].Value.ToString()), db.getRecord(Form1.OrganizedSheet[category, row].Value.ToString()), Form1.OrganizedSheet[Age, row].Value, db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString()));
            if (common.CheckEnglish(Form1.BulkSheet[9, row].Value.ToString()))
            {
                Form1.BulkSheet[9, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setCategory(int row)
        {
            Form1.BulkSheet[3, row].Value = Form1.OrganizedSheet[category, row].Value.ToString().Trim();
            Form1.BulkSheet[10, row].Value = db.getRecord(Form1.OrganizedSheet[category, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[10, row].Value.ToString()))
            {
                Form1.BulkSheet[10, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setGender(int row)
        {
            Form1.BulkSheet[4, row].Value = Form1.OrganizedSheet[Gender, row].Value.ToString();
            Form1.BulkSheet[11, row].Value = db.getRecord(Form1.OrganizedSheet[Gender, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[11, row].Value.ToString()))
            {
                Form1.BulkSheet[11, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
        private void setAge(int row)
        {
            Form1.BulkSheet[5, row].Value = getAge(Form1.OrganizedSheet[Age, row].Value.ToString())[0];
            Form1.BulkSheet[12, row].Value = db.getRecord(Form1.BulkSheet[5, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[12, row].Value.ToString()))
            {
                Form1.BulkSheet[12, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }

        }
        private void setRecommendedAge(int row)
        {
            Form1.BulkSheet[6, row].Value = getAge(Form1.OrganizedSheet[Age, row].Value.ToString())[1];
            Form1.BulkSheet[13, row].Value = db.getRecord(Form1.BulkSheet[6, row].Value.ToString());
            if (common.CheckEnglish(Form1.BulkSheet[13, row].Value.ToString()))
            {
                Form1.BulkSheet[13, row].Style.BackColor = Color.Yellow;
                UnTranslatedCount++;
            }
        }
       
        private void setLink(int row)
        {
            Form1.BulkSheet[14, row].Value = Form1.OrganizedSheet[Link, row].Value;
        }
        private void setPrice(int row)
        {
            Form1.BulkSheet[15, row].Value = Form1.OrganizedSheet[Price, row].Value;
        }
        private void setQuantity(int row)
        {
            Form1.BulkSheet[16, row].Value = Form1.OrganizedSheet[Quantity, row].Value;
        }
        private string[] getAge(string text) {
            string[] value= { "", "" };
            if (text.ToLower().Contains("all"))
            {
                value[0] = "All Ages";
                value[1] = "All Ages";
            }
            else {
                try
                {
                    if (text.Contains("-"))
                    {
                        string[] ageRange = text.Split('-');
                        int firstAge = Convert.ToInt32(ageRange[0]);
                        value[0] = string.Format("{0}  Years & above", firstAge);
                        value[1] = getRecommendedAge(firstAge).Trim();
                    }
                    else {
                        int age = Convert.ToInt32(text);
                        value[0] = string.Format("{0} Years", age);
                        value[1] = getRecommendedAge(age).Trim();
                    }
                    

                    
                }
                catch { }

            }

            return value;
        }

        private string getRecommendedAge(int age) {
            string value = "";
            if (age < 5.99)
            {
                value = "3 - 6 Years";
            } else if (age >= 6 && age < 8.99) {
                value = "6 - 9 Years";
            }
            else if (age >= 9 && age < 11.99)
            {
                value = "9 - 12 Years";
            } else if (age>=12) {
                value = "12 Years & Above";
            }

            return value.Trim();
        }

    }
}
