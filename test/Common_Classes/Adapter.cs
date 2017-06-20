using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace helper
{
    public class Adapter
    {
        public void SwitchCategory(object index)
        {

            switch (Convert.ToString(index))
            {

                case "Perfumes & Fragrances (478)":
                    Perfume perfume = new Perfume();
                    perfume.getGender();

                    break;
                case "Refrigerators & Freezers (531)":
                    Refrigerator Ref = new Refrigerator();
                    Ref.Organize();
                    break;
                case "Makeup (295)":
                    Makeup makeup = new Makeup();
                    makeup.Organize();
                    break;
                case "Mobile Phone Accessories (26)":
                    Mobile_Accessories acc = new Mobile_Accessories();
                    acc.Organize();
                    break;
                case "Tablet Accessories (181)":
                    Tablet_Accessories Tabacc = new Tablet_Accessories();
                    Tabacc.Organize();
                    break;

                case "Cable (29)":
                    Cables cable = new Cables();
                    cable.Organize();
                    break;
                case "Power Banks (562)":
                    PowerBank pBank = new PowerBank();
                    pBank.Organize();
                    break;
                case "Chargers laptop (351)":
                    Charger charger = new Charger();
                    charger.Organize();
                    break;
                case "Headphones & Headsets (373)":
                    HeadSet headSet = new HeadSet();
                    headSet.Organize();
                    break;
                case "Watches (490)":
                    Watches watches = new Watches();
                    watches.Organize();
                    break;
                case "Tops (488)":
                    Tops tops = new Tops();
                    tops.Organize();
                    break;
                case "Casual & Dress Shoes (481)":
                    Casual_Shoes shoes = new Casual_Shoes();
                    shoes.Organize();
                    break;
                case "Pants (477)":
                    Pants pants = new Pants();
                    pants.Organize();
                    break;
                case "Slippers (485)":
                    Slipper slipper = new Slipper();
                    slipper.Organize();
                    break;
                case "Handbags (472)":
                    Handbags handbag = new Handbags();
                    handbag.Organize();
                    break;
                case "Power Tools (97)":
                    Power_Tool power = new Power_Tool();
                    power.Organize();
                    break;
                case "Hand Tools (319)":
                    Hand_Tool hand = new Hand_Tool();
                    hand.Organize();
                    break;

                case "Baby Clothes (343)":
                    Baby_Clothes baby = new Baby_Clothes();
                    baby.Organize();

                    break;
                case "Sleepwear (484)":
                    Sleepwear jack = new Sleepwear();
                    jack.Organize();
                    break;


                default:
                    MessageBox.Show("Please choose category first");
                    break;
            }


            Form1.LogWrite("Finished Analayzing category \"" + Convert.ToString(index) + "\"");
            Form1.LogActionsTotal = 0;


        }


        public void SwitchBulk(object index)
        {

            switch (Convert.ToString(index))
            {

                case "Perfumes & Fragrances (478)":

                    Perfume per = new Perfume();

                    per.createBulk();


                    break;
                case "Refrigerators & Freezers (531)":
                    Refrigerator Ref = new Refrigerator();
                    Ref.creatBulk();
                    break;
                case "Makeup (295)":
                    Makeup makeup = new Makeup();
                    makeup.createBulk();
                    break;
                case "Mobile Phone Accessories (26)":
                    Mobile_Accessories acc = new Mobile_Accessories();
                    acc.createBulk();
                    break;
                case "Tablet Accessories (181)":
                    Tablet_Accessories Tabacc = new Tablet_Accessories();
                    Tabacc.createBulk();
                    break;


                case "Cable (29)":
                    Cables cable = new Cables();
                    cable.createBulk();
                    break;
                case "Power Banks (562)":
                    PowerBank pBank = new PowerBank();
                    pBank.createBulk();
                    break;
                case "Chargers laptop (351)":
                    Charger charger = new Charger();
                    charger.createBulk();
                    break;

                case "Headphones & Headsets (373)":
                    HeadSet headSet = new HeadSet();
                    headSet.createBulk();
                    break;
                case "Watches (490)":
                    Watches watches = new Watches();
                    watches.createBulk();
                    break;
                case "Tops (488)":
                    Tops tops = new Tops();
                    tops.createBulk();
                    break;
                case "Casual & Dress Shoes (481)":
                    Casual_Shoes shoes = new Casual_Shoes();
                    shoes.createBulk();
                    break;

                case "Pants (477)":
                    Pants pants = new Pants();
                    pants.createBulk();
                    break;
                case "Slippers (485)":
                    Slipper slipper = new Slipper();
                    slipper.createBulk();
                    break;
                case "Handbags (472)":
                    Handbags handbag = new Handbags();
                    handbag.createBulk();
                    break;

                case "Power Tools (97)":
                    Power_Tool power = new Power_Tool();
                    power.createBulk();
                    break;
                case "Hand Tools (319)":
                    Hand_Tool hand = new Hand_Tool();
                    hand.createBulk();
                    break;
                case "Baby Clothes (343)":
                    Baby_Clothes baby = new Baby_Clothes();
                    baby.createBulk();
                    break;
                case "Sleepwear (484)":
                    Sleepwear jac = new Sleepwear();
                    jac.createBulk();
                    break;

                default:
                    MessageBox.Show("Please choose category first");
                    break;
            }
            Form1.LogWrite(string.Format("Finished creating bulk for category {0}", Convert.ToString(index)));

        }


    }


}
