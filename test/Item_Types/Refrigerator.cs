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
        int Brand = 1, Model = 2, Type = 3, Capacity = 4, Color = 5, Des = 6, Dims = 7, Material = 8;
        bool isFirst = true;
        public void Organize()
        {
            int row, col;
            Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.RowCount = Form1.OrganizedSheet.RowCount));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[0].HeaderText = "Row Number"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[Brand].HeaderText = "Brand"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[Model].HeaderText = "Model Number"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[Type].HeaderText = "Type"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[Capacity].HeaderText = "Capacity"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[Color].HeaderText = "Color"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[Des].HeaderText = "Description"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[Dims].HeaderText = "Dimensions(Depth x Height x Width"));
            Form1.OrganizedSheet.Invoke(new Action(() => Form1.OrganizedSheet.Columns[Material].HeaderText = "Material"));
            for (row = 0; row < Form1.Sheet.RowCount; row++)
            {
                for (col = 0; col < Form1.Sheet.ColumnCount; col++)
                {
                    try
                    {
                        Form1.OrganizedSheet[0,row].Value = Convert.ToString(row + 2);
                    }
                    catch { }
                    try
                    {
                        Regex regexSize = new Regex(@"([Rr]efr?\w+)|(\w+eezer)");
                        Match matchSize = regexSize.Match(Form1.Sheet[col,row].Value.ToString());



                        if (matchSize.Success)
                        {
                            Form1.OrganizedSheet[Type,row].Value = matchSize.Value;

                        }
                    }
                    catch { }
                }
            }
        }

        public void creatBulk()
        {
            if (isFirst)
            {

                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[0].HeaderText = "Title"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Bran", "Brand")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Des", "Description")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Color", "Color")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Depth", "Depth")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Height", "Height")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Width", "Width")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Material", "Material")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("model", "Model Number")));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns.Add("Capacity", "Capacity")));

                isFirst = false;
            }
            else
            {
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[0].HeaderText = "Title"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[1].HeaderText = "Brand"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[2].HeaderText = "Description"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[3].HeaderText = "Color"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[4].HeaderText = "Depth"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[5].HeaderText = "Height"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[6].HeaderText = "Width"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[7].HeaderText = "Material"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[8].HeaderText = "Model Number"));
                Form1.BulkSheet.Invoke(new Action(() => Form1.BulkSheet.Columns[9].HeaderText = "Capacity"));
            }
            for (int i = 0; i < Form1.OrganizedSheet.RowCount; i++)
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
        }



        public void setTitle(int row)
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            string title = Form1.OrganizedSheet[Brand, row].Value + " " + Form1.OrganizedSheet[Model, row].Value + " " + Form1.OrganizedSheet[Type, row].Value + " - "
                + Form1.OrganizedSheet[Capacity, row].Value + ", " + Form1.OrganizedSheet[Color, row].Value;
            Form1.BulkSheet[0, row].Value = textInfo.ToTitleCase(title);
        }
        public void setBrand(int row)
        {
            Form1.BulkSheet[1, row].Value = Form1.OrganizedSheet[Brand, row].Value;
        }
        public void setColor(int row)
        {
            Form1.BulkSheet[3, row].Value = Form1.OrganizedSheet[Color, row].Value;
        }
        public void setModel(int row)
        {
            Form1.BulkSheet[8, row].Value = Form1.OrganizedSheet[Model, row].Value;
        }

        public void setCapacity(int row)
        {
            Form1.BulkSheet[9, row].Value = Form1.OrganizedSheet[Capacity, row].Value;
        }

        public void SetMaterial(int row)
        {
            Form1.BulkSheet[7, row].Value = Form1.OrganizedSheet[Material, row].Value;
        }
        public void setDescription(int row)
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
        public void setDimensions(int row)
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

    }
}
