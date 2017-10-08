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
                case "Kitchen and Dining (316)":
                    Kitchen_Dinning Kitchen = new Kitchen_Dinning();
                    Kitchen.Organize();
                    break;
                case "Underwear (489)":
                    Underwears under = new Underwears();
                    under.Organize();
                    break;
                case "Snacks & Bakery (582)":
                    Snacks_Bakery snack = new Snacks_Bakery();
                    snack.Organize();
                    break;
                case "Swimwear (487)":
                    Swimwear swim = new Swimwear();
                    swim.Organize();
                    break;
                case "Toys (24)":
                    Toys toys = new Toys();
                    toys.Organize();
                    break;
                case "Baby Gears (331)":
                    Baby_Gear baby_gear = new Baby_Gear();
                    baby_gear.Organize();
                    break;
                case "Baby Safety and Health (332)":
                    Baby_Safety baby_safety = new Baby_Safety();
                    baby_safety.Organize();
                    break;
                case "Baby Accessories (430)":
                    Baby_Accessories baby_acces = new Baby_Accessories();
                    baby_acces.Organize();
                    break;

                case "Rings (284)":
                    Rings rings = new Rings();
                    rings.Organize();
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

                case "Kitchen and Dining (316)":
                    Kitchen_Dinning kitchen = new Kitchen_Dinning();
                    kitchen.createBulk();
                    break;
                case "Underwear (489)":
                    Underwears under = new Underwears();
                    under.createBulk();
                    break;
                case "Snacks & Bakery (582)":
                    Snacks_Bakery snack = new Snacks_Bakery();
                    snack.createBulk();
                    break;
                case "Swimwear (487)":
                    Swimwear swim = new Swimwear();
                    swim.createBulk();
                    break;
                case "Toys (24)":
                    Toys toys = new Toys();
                    toys.createBulk();
                    break;
                case "Baby Gears (331)":
                    Baby_Gear baby_gear = new Baby_Gear();
                    baby_gear.createBulk();
                    break;
                case "Baby Safety and Health (332)":
                    Baby_Safety baby_safety = new Baby_Safety();
                    baby_safety.createBulk();
                    break;
                case "Baby Accessories (430)":
                    Baby_Accessories baby_acces = new Baby_Accessories();
                    baby_acces.createBulk();
                    break;
                case "Rings (284)":
                    Rings rings = new Rings();
                    rings.createBulk();
                    break;

                default:
                    MessageBox.Show("Please choose category first");
                    break;
            }
            Form1.LogWrite(string.Format("Finished creating bulk for category {0}", Convert.ToString(index)));

        }


    }


}
