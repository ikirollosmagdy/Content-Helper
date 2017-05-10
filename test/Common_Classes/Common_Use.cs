using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace helper
{
  public  class Common_Use
    {
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
    }
}
