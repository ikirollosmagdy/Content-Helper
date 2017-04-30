using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;

namespace helper
{
    public partial class Updater : Form
    {
        string baseUrl;
        public  WebClient downloader;
        List<string> FileNames = new List<string>();
        public Updater(string baseUrl)
        {
            InitializeComponent();
            this.baseUrl = baseUrl;
        }

        private void Updater_Load(object sender, EventArgs e)
        {
           
           

        }




        private void Updater_Activated(object sender, EventArgs e)
        {
            btnFinish.Enabled = false;


            XDocument document = XDocument.Load(baseUrl + "UpdateInfo.xml");
            var elements = document.Elements("AppName").Elements();
            int elementsCount = elements.Count();
            for (int i = 0; i < elementsCount; i++)
            {
                FileNames.Add(elements.ElementAt(i).Value);

            }
            ProgressTotal.Maximum = FileNames.Count-1;
            for (int x = 1; x < FileNames.Count; x++)
            {
          DownLoadFile(baseUrl + FileNames[x], FileNames[x].Replace('!',' '));
            }

        }
        private void DownLoadFile(string source, string target)
        {


            
            downloader = new WebClient();
            // fake as if you are a browser making the request.
            downloader.Headers.Add("User-Agent", "Mozilla/4.0 (compatible; MSIE 8.0)");
            downloader.DownloadFileCompleted += new AsyncCompletedEventHandler(Downloader_DownloadFileCompleted);
            downloader.DownloadProgressChanged +=
                new DownloadProgressChangedEventHandler(Downloader_DownloadProgressChanged);
            downloader.DownloadFileAsync(new Uri(source), target);
            // wait for the current thread to complete, since the an async action will be on a new thread.
            while (downloader.IsBusy) { }
        }
             private void Downloader_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            ProgressFile.Value = e.ProgressPercentage;
            lblDownloading.Text = string.Format("&Downloaded {0} %", e.ProgressPercentage);
        }

        private void Downloader_DownloadFileCompleted(object sender, AsyncCompletedEventArgs e)
        {
            // display completion status.
            if (e.Error != null)
                Console.WriteLine(e.Error.Message);
            else
            {
                ProgressTotal.PerformStep();
                if (ProgressTotal.Value == ProgressTotal.Maximum)
                {
                    btnFinish.Enabled = true;
                }

            }
           
        }

        private void btnFinish_Click(object sender, EventArgs e)
        {
            Application.Restart();
        }
    }


    
  
}
