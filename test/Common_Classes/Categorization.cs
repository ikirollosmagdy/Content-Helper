using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace helper
{
 public   class Categorization
    {
        public string getCategory(string text)
        {
            string cat="";
            string pat = @"(Adapter|Cable|Splitter|Power cord)|(case|frame|Band|Cover|Clip|Flash|Handset|Jack|Lens Protector|Maintenance Set|Holder|Mount|Remote Shutter|Replacement Part|Screen Magnifier|Security Accessory|Selfie|SIM|Stand|Strap|Stylus|Tripod)";

            // Instantiate the regular expression object.
            Regex r = new Regex(pat, RegexOptions.IgnoreCase);

            // Match the regular expression pattern against a text string.
            Match m = r.Match(text);

            while (m.Success)
            {

                for (int i = 1; i <= m.Groups.Count; i++)
                {
                    if (m.Groups[i].Value != string.Empty)
                    {
                        switch (i)
                        {
                            case 1:
                                cat = "Cable";
                                break;
                            case 2:
                                cat = "Mobile Phone Accessories";
                                break;
                           
                        }
                    }



                }
                m = m.NextMatch();
            }

            return cat;

        }
    }
}
