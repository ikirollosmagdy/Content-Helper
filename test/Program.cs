using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace helper
{
    static class Program
    {
       static int errorCount = 0;
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.ThreadException +=
               new ThreadExceptionEventHandler(Application_ThreadException);
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
           
        }

        static void Application_ThreadException(object sender, ThreadExceptionEventArgs e)
        {
            errorCount++;
            using (StreamWriter sw = File.AppendText("CrashLog.txt"))
            {
                sw.WriteLine(string.Format("============================== Error #{0} Start ============================", errorCount));
                sw.WriteLine(e.Exception.ToString());
                
                sw.WriteLine(string.Format("============================== Error #{0} End =============================={1}", errorCount,Environment.NewLine));
                sw.Close();
            }

        }
    }
}
