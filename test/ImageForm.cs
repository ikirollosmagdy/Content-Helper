using System;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;

namespace helper
{
    public partial class ImageForm : Form
    {
       public static string Urls;
         string[] arrUrls;
        int imagecount = -1;
        public ImageForm(string url)
        {
            InitializeComponent();
            Urls = url;
            arrUrls= url.Split('\n');
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
            string ext = Path.GetExtension(arrUrls[imagecount]);

            progressBar1.Hide();
           
            lblSize.Text = string.Format("Size: {0} * {1} with Format: {2}",width,height,ext);
        
            if (height<400 || width < 400|| height>600||width>600)
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
            imagecount++;
            pictureBox1.LoadAsync(arrUrls[imagecount]);
        }

        private void toolStripStatusLabel3_Click(object sender, EventArgs e)
        {
            if (imagecount == arrUrls.Length-1)
            {
                return;
            }
            imagecount++;
            progressBar1.Show();
            pictureBox1.LoadAsync(arrUrls[imagecount]);
            
        }

        private void toolStripStatusLabel4_Click(object sender, EventArgs e)
        {
            if (imagecount == 0)
            {
                return;
            }
            imagecount--;
            progressBar1.Show();
            pictureBox1.LoadAsync(arrUrls[imagecount]);
           
        }
    }
}
