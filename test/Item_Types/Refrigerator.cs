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
    public class Refrigerator


    {
        int Brand = 1, Model = 2, Type = 3, Capacity = 4, Color = 5, Des = 6, Dims = 7, Material = 8
            ,Link=9,Price=10,Quantity=11;
       

        private void setupTable()
        {
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Clear()));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("No", "Row Number")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Brand", "Brand")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Model", "Model")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Type", "Type")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Capacity", "Capacity")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Color", "Color")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Des", "Desciption")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Dims", "Dimensions(Depth x Height x Width")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Material", "Material")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Link", "Link")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Price", "Price")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns.Add("Quantity", "Quantity")));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.RowCount = Form1.Sheet.RowCount));
            foreach(DataGridViewColumn co in Form1.OrganizedSheet.Columns)
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
                        Form1.OrganizedSheet[0, row].Value = Convert.ToString(row + 2);
                    }
                    catch { }
                    try
                    {
                        Regex regexSize = new Regex(@"([Rr]efr?\w+)|(\w+eezer)");
                        Match matchSize = regexSize.Match(Form1.Sheet[col, row].Value.ToString());
                          if (matchSize.Success)
                        {
                            Form1.OrganizedSheet[Type, row].Value = matchSize.Value;

                        }
                    }
                    catch { }
                    try
                    {
                        Regex regexModel = new Regex(@"(?!\s+)((\w+([-|/| ])?)?\w+([-|/| ])?\d+(\w+|\d+)?([-|/| ])?(\w+|\d+)?)");
                        Match matchModel = regexModel.Match(Form1.Sheet[col, row].Value.ToString());
                        
                        if (matchModel.Success)
                        {
                            Form1.OrganizedSheet[Model, row].Value = matchModel.Value;

                        }
                    }
                    catch { }
                }
            }
        }

        public void creatBulk()
        {

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
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.RowCount = Form1.OrganizedSheet.RowCount));

            for (int i = 0; i < Form1.OrganizedSheet.RowCount; i++)
            {
                try
                {
                    setTitle(i);
                    setBrand(i);
                    setDescription(i);
                    setColor(i);
                    setDimensions(i);
                    SetMaterial(i);
                    setModel(i);
                    setCapacity(i);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }


            }
        }



         void setTitle(int row)
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            string title = Form1.OrganizedSheet[Brand, row].Value + " " + Form1.OrganizedSheet[Model, row].Value + " " + Form1.OrganizedSheet[Type, row].Value + " - "
                + Form1.OrganizedSheet[Capacity, row].Value + ", " + Form1.OrganizedSheet[Color, row].Value;
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase(title);
        }
         void setBrand(int row)
        {
            Form1.BulkSheet[1, row].Value = Form1.OrganizedSheet[Brand, row].Value;
        }
         void setColor(int row)
        {
            Form1.BulkSheet[3, row].Value = Form1.OrganizedSheet[Color, row].Value;
        }
         void setModel(int row)
        {
            Form1.BulkSheet[8, row].Value = Form1.OrganizedSheet[Model, row].Value;
        }

         void setCapacity(int row)
        {
            Form1.BulkSheet[9, row].Value = Form1.OrganizedSheet[Capacity, row].Value;
        }

         void SetMaterial(int row)
        {
            Form1.BulkSheet[7, row].Value = Form1.OrganizedSheet[Material, row].Value;
        }
         void setDescription(int row)
        {
            string des = "";
            try
            {

                string[] lines = Form1.OrganizedSheet[Des, row].Value.ToString().Split('\n');
                foreach (string line in lines)
                {
                    des = des + "<p>" + line + "</p>";
                }
            }
            catch { }
            des = des + "<li>Color: " + Form1.OrganizedSheet[Color, row].Value + "</li><li>Dimensions: " + Form1.OrganizedSheet[Dims, row].Value +
"</li><li>Capacity: " + Form1.OrganizedSheet[Capacity, row].Value + "</li>";

            Form1.BulkSheet[2, row].Value = des;
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
                }
                catch { }
            }

        }

    }
}
