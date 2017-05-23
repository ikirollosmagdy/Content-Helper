using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
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

            Thread t = new Thread(()=> { 
            try
            {
                string[] Scopes = { SheetsService.Scope.Spreadsheets };
                string ApplicationName = "Google Sheets API .NET Quickstart";

                UserCredential credential;

                using (var stream =
                    new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
                {
                    string credPath = System.Environment.GetFolderPath(
                        System.Environment.SpecialFolder.Personal);


                    credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,
                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
                    Console.WriteLine("Credential file saved to: " + credPath);
                }

                // Create Google Sheets API service.
                var service = new SheetsService(new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = ApplicationName,
                });



                String spreadsheetId2 = "1AxuO473BEWpVqOtR0wfOOPPVrhUC1ta_Xxj9i79seXA";
                String range2 = "CrashLog!A:Z";

                ValueRange valueRange = new ValueRange();
                valueRange.MajorDimension = "ROWS";//"ROWS";//COLUMNS
             
              
              
                var oblist = new List<object>() {
                    Environment.UserName,e.Exception.ToString()

                };
                valueRange.Values = new List<IList<object>> { oblist };

                // SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId2, range2);
                SpreadsheetsResource.ValuesResource.AppendRequest update = service.Spreadsheets.Values.Append(valueRange, spreadsheetId2, range2);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                AppendValuesResponse result2 = update.Execute();


                // End of posting and Close app
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            });
            t.Start();

        }
    }
}
