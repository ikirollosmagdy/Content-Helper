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
           
           
        }

        private void ImageForm_Load(object sender, EventArgs e)
        {

        }

        private void pictureBox1_LoadCompleted(object sender, AsyncCompletedEventArgs e)
        {
            int height = pictureBox1.Image.Height,
              width = pictureBox1.Image.Height;


            progressBar1.Hide();
            lblSize.Text = width.ToString() + "*" + height.ToString();
        
            if (height<400 || width < 400)
            {
             //   pictureBox2.Image = Properties.Resources.Bad;
                lblComment.Text="We suggest to change this image";
            }
            else
            {
              //  pictureBox2.Image = Properties.Resources.Good;
                lblComment.Text = "Good to be in SOUQ.com";
            }

        }

        private void ImageForm_Shown(object sender, EventArgs e)
        {
            progressBar1.Show();
            
            pictureBox1.LoadAsync(Url);
        }
    }
}
