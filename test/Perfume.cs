using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace test
{
    class Perfume

    {



        public static int Brand = 1, Gender = 5, Size = 4, Type = 3, PerfumeName = 2,
            ExtraData = 6,FregFamily=7;
      static  bool  isFirst = true;

        public static void createBulk()
        {

           

            if (isFirst)
                {
                //  Form1.BulkSheet.Columns.Add("Title", "Title");
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[0].HeaderText="Title"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Bran", "Brand")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Des", "Description")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Type", "Fregrance Type")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Size", "Size")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Gender", "Target Group")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Familey", "Fragrance Family")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("FragName", "Perfume Name")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("arType", "Fregrance Type AR")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Sizear", "Size Ar")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("GenderAr", "Target Group AR")));
                    isFirst = false;
                }
                else
                {
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[0].HeaderText = "Title"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[1].HeaderText="Brand"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[2].HeaderText="Description"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[3].HeaderText = "Fregrance Type"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[4].HeaderText = "Size"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[5].HeaderText = "Target Group"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[6].HeaderText = "Fragrance Family"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[7].HeaderText = "Perfume Name"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[8].HeaderText = "Fregrance Type AR"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[9].HeaderText = "Size AR"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[10].HeaderText = "Target Group AR"));
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
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[Type].HeaderText = "Fregrance type"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[ExtraData].HeaderText = "Description"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[FregFamily].HeaderText = "Fregrance Family"));
            for (row = 0; row < Form1.Sheet.RowCount; row++)
            {
                for (col = 0; col < Form1.Sheet.RowCount; col++)
                {
                    try
                    {
                        Regex regexSize = new Regex(@"(\d{2,3} ?[Mm]l)");
                        Match matchSize = regexSize.Match(Form1.Sheet.Rows[row].Cells[col].Value.ToString());



                        if (matchSize.Success)
                        {
                            Form1.OrganizedSheet.Rows[row].Cells[Size].Value = matchSize.Value;
                        }
                    }
                    catch (Exception) { }
                    try
                    {
                        Regex regexGender = new Regex(@"(([Ww]o)?[Mm][ae]n)|([Uu]ni(sex)?)|([Bb]oys?)|([Gg]irls?)");
                        MatchCollection matchGender = regexGender.Matches(Form1.Sheet.Rows[row].Cells[col].Value.ToString());
                        if (matchGender.Count > 1)
                        {
                            Form1.OrganizedSheet.Rows[row].Cells[Gender].Value = matchGender[0].Value + ", " + matchGender[1].Value;
                        }
                        else
                        {
                            Form1.OrganizedSheet.Rows[row].Cells[Gender].Value = matchGender[0].Value;
                        }

                    }
                    catch (Exception) { }
                    try
                    {
                        Regex regType = new Regex(@"((?=[Ee]au)( ?\w+ ?\w+ ?\w{5,8}))|((?=[Pp]erf)(\w{6} ?\w+ ?\w{3,6}))|(([A-Z]\w+ ?\w+ ?)(?=[Pp]arf)\w{6})|([Oo]ud ?)|([Ee][Dd][TtPpCc] ?)");
                        Match matchType = regType.Match(Form1.Sheet.Rows[row].Cells[col].Value.ToString());
                        if (matchType.Success)
                        {
                            Form1.OrganizedSheet.Rows[row].Cells[Type].Value = matchType.Value;

                        }
                    }
                    catch (Exception ) { }
                    try
                    {
                        Form1.OrganizedSheet.Rows[row].Cells[0].Value = Convert.ToString(row + 2);
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                }
            }
            
        }




        public static void setTitle(int row)
        {

            
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
             

            string title = Form1.OrganizedSheet.Rows[row].Cells[PerfumeName].Value + " By " + Form1.OrganizedSheet.Rows[row].Cells[Brand].Value
              + " For " + Form1.OrganizedSheet.Rows[row].Cells[Gender].Value + " - " + Form1.OrganizedSheet.Rows[row].Cells[Type].Value
               + ", " + Form1.OrganizedSheet.Rows[row].Cells[Size].Value;
            Form1.BulkSheet.Rows[row].Cells[0].Value = textInfo.ToTitleCase(title);

        }

        public static void setDescription(int row)
        {

            string Des = "<p>" + Form1.OrganizedSheet.Rows[row].Cells[ExtraData].Value
                + "</p> <p><strong>Product Features:</strong></p> <ul> <li>Brand: "
                + Form1.OrganizedSheet.Rows[row].Cells[Brand].Value + "</li> <li>Target Group: "
                + Form1.OrganizedSheet.Rows[row].Cells[Gender].Value + "</li> <li>Fragrance Type: "
                + Form1.OrganizedSheet.Rows[row].Cells[Type].Value + "</li> <li>Perfume Name: "
                + Form1.OrganizedSheet.Rows[row].Cells[PerfumeName].Value + "</li> <li>Size: "
                + Form1.OrganizedSheet.Rows[row].Cells[Size].Value + "</li> </ul>";
            Form1.BulkSheet.Rows[row].Cells[2].Value = Des;


        }
        public static void setBrand(int row)
        {
            Form1.BulkSheet.Rows[row].Cells[1].Value = Form1.OrganizedSheet.Rows[row].Cells[Brand].Value;
        }
        public static void setType(int row)
        {
            Form1.BulkSheet.Rows[row].Cells[3].Value = Form1.OrganizedSheet.Rows[row].Cells[Type].Value;
           
                if (Form1.OrganizedSheet.Rows[row].Cells[Type].Value != null)
                {
                    
                     Form1.BulkSheet.Rows[row].Cells[8].Value = getArabic(Form1.BulkSheet.Rows[row].Cells[3].Value.ToString());
                    

                  
                
            }
        }
        public static void setSize(int row)
        {
            Form1.BulkSheet.Rows[row].Cells[4].Value = Form1.OrganizedSheet.Rows[row].Cells[Size].Value;
            if (Form1.OrganizedSheet.Rows[row].Cells[Size].Value != null)
            {
                Form1.BulkSheet.Rows[row].Cells[9].Value = getArabic(Form1.BulkSheet.Rows[row].Cells[4].Value.ToString());
            }

        }
        public static void setGender(int row)
        {
            Form1.BulkSheet.Rows[row].Cells[5].Value = Form1.OrganizedSheet.Rows[row].Cells[Gender].Value;

            if (Form1.OrganizedSheet.Rows[row].Cells[Gender].Value != null)
            {
               Form1.BulkSheet.Rows[row].Cells[10].Value = getArabic(Form1.BulkSheet.Rows[row].Cells[5].Value.ToString());
            }
           
        }

        public static void setFragFamiley(int row)
        {
            Form1.BulkSheet.Rows[row].Cells[6].Value = Form1.OrganizedSheet.Rows[row].Cells[FregFamily].Value;
        }
        public static void setPerfumeName(int row)
        {
            Form1.BulkSheet.Rows[row].Cells[7].Value = Form1.OrganizedSheet.Rows[row].Cells[PerfumeName].Value;
        }



        public static string getArabic(String Text)
        {
            string arabicMatch = Text;
            if (Text != null)
            {
                try
                {
                    foreach (string line in System.IO.File.ReadAllLines("lookup.dat"))
                    {
                        if (line.Contains(Text))
                            arabicMatch = line.Split('	')[1];
                    }
                }
                catch (Exception )
                {
                    arabicMatch="";
                }
            }
            return arabicMatch;
        }


    }
}
