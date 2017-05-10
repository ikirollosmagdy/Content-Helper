using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace helper
{
  public  class Common_Use
    {
        ///<summary>
        ///Gets the replacement for words like family colors
        ///</summary>
        public string getReplacement(string Text)
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            string arabicMatch = Text;
            if (Text != null)
            {
                Text = Text.ToLower().Trim();
                try
                {
                    foreach (string line in System.IO.File.ReadAllLines("lookup.dat"))
                    {

                        if (line.Contains(Text))
                            arabicMatch = line.Split('	')[1];

                    }
                }
                catch (Exception)
                {
                    arabicMatch = "";
                  
                }
            }
            return textInfo.ToTitleCase(arabicMatch);
        }

        ///<summary>
        ///Checks if the cell has untranslated words
        ///</summary>
        public bool CheckEnglish(string text)
        {
            bool IsEnglish = false;
            Regex regex = new Regex(@"[^pulbi<>\/\d\.,\s]([a-zA-Z])");
            Match match = regex.Match(text);
            if (match.Success)
            {
                IsEnglish = true;
            }
            else
            {
                IsEnglish = false;
            }

            return IsEnglish;
        }
        ///<summary>
        ///Send The Offline report to google sheet after Internet connection
        ///</summary>
        public void sendOfflineReport()
        {
            if (File.Exists("OfflineReport.dat"))
            {
                string[] Lines;
                foreach (string line in System.IO.File.ReadAllLines("OfflineReport.dat"))
                {
                    Lines = line.Split('\t').ToArray();
                    Console.WriteLine("Offline Data array"+Lines);
                    // Posting data to google Spreadsheet 
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
                        String range2 = "Sheet1!A:Z";

                        ValueRange valueRange = new ValueRange();
                        valueRange.MajorDimension = "ROWS";//"ROWS";//COLUMNS

                        var oblist = new List<object>() {
                    Lines[0],Convert.ToInt32( Lines[1]),
                 Convert.ToInt32( Lines[2]),
                    Lines[3],
                   Lines[4],
                    Lines[5],
                   Lines[6],
                     Lines[7],
                    Lines[8],
                  Convert.ToInt32(Lines[9]),
                    Lines[10],Lines[11]
                };
                        valueRange.Values = new List<IList<object>> { oblist };

                        // SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId2, range2);
                        SpreadsheetsResource.ValuesResource.AppendRequest update =  service.Spreadsheets.Values.Append(valueRange, spreadsheetId2, range2);
                        update.ValueInputOption = SpreadsheetsResource.ValuesResource.AppendRequest.ValueInputOptionEnum.RAW;
                       AppendValuesResponse result2 =  update.Execute();

                      
                        // End of posting and Close app
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error offline data"+ex.Message);
                    }
                    File.Delete("OfflineReport.dat");
                }
            }
           
        }


        ///<summary>
        ///Translate color cell in sperate words from database 
        ///</summary>
        public string getColorMulti(string text)
        {
            Database db = new Database();
            string arColor = "";
            string[] colorAr = text.Split(' ');
            if (colorAr.Length > 1)
            {
                arColor = db.getRecord(colorAr[1]) + " " + db.getRecord(colorAr[0]);
            }
            else
            {
                arColor = db.getRecord(text);
            }
            return arColor;
        }
    }
}
