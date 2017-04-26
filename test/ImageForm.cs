using System;
using System.ComponentModel;
using System.Windows.Forms;

namespace helper
{
    public partial class ImageForm : Form
    {
       public string Url;
        public ImageForm()
        {
            InitializeComponent();
        }

        private void ImageForm_Activated(object sender, EventArgs e)
        {
            progressBar1.Show();
         
            
            pictureBox1.LoadAsync(Url);
           
        }

        private void ImageForm_Load(object sender, EventArgs e)
        {

        }

        private void pictureBox1_LoadCompleted(object sender, AsyncCompletedEventArgs e)
        {
            int height = pictureBox1.Image.Height,
              width = pictureBox1.Image.Height;


            progressBar1.Hide();
            lblHeight.Text = height.ToString();
            lblWidth.Text = width.ToString();
            if (height<400 || width < 400)
            {
                pictureBox2.Image = Properties.Resources.Bad;
                lblSouq.Text="We suggest to change this image";
            }
            else
            {
                pictureBox2.Image = Properties.Resources.Good;
                lblSouq.Text = "Good to be in SOUQ.com";
            }

        }
    }
}
