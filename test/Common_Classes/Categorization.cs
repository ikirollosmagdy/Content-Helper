using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace helper
{
    public class Categorization
    {
        public string getCategory(string text)
        {
            string cat = "";
            string pat = @"(Adapter|Cable|Splitter|Power cord)|(case|frame|Band|Cover|Clip|Flash|Handset|Jack|Lens Protector|
Maintenance Set|Holder|Mount|Remote Shutter|Replacement Part|Screen Magnifier|Security Accessory|Selfie|SIM|Stand|Strap|Stylus|
Tripod)|(ballerina|espadri\w+|sneaker|flat|heel\w+|wedg\w+)|(briefcase|hobo|clutch)|
(legg?\w+|jeans|trous\w+)|(clog|flip ?flops?|slides?|slippers?|moccasin)|(ear)|
(Analog|Watch|Digital)|
(Blotting Paper|Blusher|Concealer|Foundation|Powder|Bronzers|Eyebrow|Eyelash|Eyeliner|Eyeshadow|Foundation|Face Powder|Face Primers|Glitter|Contour|Lip|Makeup|Mascara|Nail Polish)|
(Alternative Ice Cubes|Apron|Bamboo Skewer|Beverage Decoration Templates|Bottle & Cups Sleeve Holder|Chopstick|Shaker|Cookware Holder|Cups Holder|Drink Chillers|Egg Mold|Egg Slicer|Food & Sauce Container|Food Bags Holder|Gloves|Holder|Hot Dog Slicer|Ice Bucket|Ice Tray|Jar|Kitchen Bag Clips|Kitchen Lighter|Kitchen Tissues Holder|Knives Holder|Liquid Dispenser Bottle|Lunch Box|Microwave Lid|Oil Sprayer|Pot Holder|Refrigerator Magnets|Ring|Safe Slice Finger Protector|Shrimp Peeler|Spice Rack|Spoons Holder|Straw|Sushi Maker|Sushi Mat|Toothpicks|Water Filter|Water Filters Replacement|Wine Pourer|Wine Stopper|
([Rr]efr?\w+|\w+eezer))|
(Stapler|Sprayer|axe|Brush|Chisel|Punch|Clamp|Vice|Cutter|Gripper|Hammer|Plier|Riveter|SandPaper|Scissor|Screwdriver|Wrench)|
(polisher|saw|cutter|blower|drill|grinder|planer|screwdriver|generator|sander|nailer|stapler)";

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
                            case 3:
                                cat = "Casual Shoes";
                                break;
                            case 4:
                                cat = "Hand bag";
                                break;
                            case 5:
                                cat = "Pants";
                                break;
                            case 6:
                                cat = "Slipper";
                                break;
                            case 7:
                                cat = "Headset";
                                break;
                            case 8:
                                cat = "Watches";
                                break;
                            case 9:
                                cat = "Makeup";
                                break;
                            case 10:
                                cat = "Kitchen & Dinning";
                                break;
                            case 11:
                                cat = "Refrigerators";
                                break;
                            case 12:
                                cat = "Hand Tool";
                                break;
                            case 13:
                                cat = "Power Tool";
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
