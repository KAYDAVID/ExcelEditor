using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ListS
{
    public partial class ColumnsSettingsForm_ListS : Form
    {
        public void DisplayCurrentSettings ()
        {
            //Column Name Display Setup
            firstNameTextBox.Text = Properties.Settings.Default.FirstNameSetting; // Default ATTENDED 0
            secondNameTextBox.Text = Properties.Settings.Default.SecondNameSetting; // Default FIRST NAME 2
            thirdNameTextBox.Text = Properties.Settings.Default.ThirdNameSetting; // Default LAST NAME 3
            fourthNameTextBox.Text = Properties.Settings.Default.FourthNameSetting;  // Default STATUS 8

            //Column Number Display Setup
            firstNumSelect.Value = Properties.Settings.Default.FirstNumSetting; // Default Column 99          
            secondNumSelect.Value = Properties.Settings.Default.SecondNumSetting; // Default Column 2
            thirdNumSelect.Value = Properties.Settings.Default.ThirdNumSetting; // Default Column 3
            fourthNumSelect.Value = Properties.Settings.Default.FourthNumSetting; // Default Column 8
        }

        public void SaveSettings ()
        {
            //Column Name Display Setup
            Properties.Settings.Default["FirstNameSetting"] = firstNameTextBox.Text;
            Properties.Settings.Default["SecondNameSetting"] = secondNameTextBox.Text;
            Properties.Settings.Default["ThirdNameSetting"] = thirdNameTextBox.Text;
            Properties.Settings.Default["FourthNameSetting"] = fourthNameTextBox.Text;

            //Column Number Display Setup
            Properties.Settings.Default["FirstNumSetting"] = firstNumSelect.Value;
            Properties.Settings.Default["SecondNumSetting"] = secondNumSelect.Value;
            Properties.Settings.Default["ThirdNumSetting"] = thirdNumSelect.Value;
            Properties.Settings.Default["FourthNumSetting"] = fourthNumSelect.Value;
        }

        public void RestoreDefaults()
        {
            //Column Name Display Setup
            Properties.Settings.Default["FirstNameSetting"] = "ATTENDED";
            Properties.Settings.Default["SecondNameSetting"] = "FIRST NAME";
            Properties.Settings.Default["ThirdNameSetting"] = "LAST NAME";
            Properties.Settings.Default["FourthNameSetting"] = "STATUS";

            //Column Number Display Setup
            Properties.Settings.Default["FirstNumSetting"] = Convert.ToDecimal(99);
            Properties.Settings.Default["SecondNumSetting"] = Convert.ToDecimal(2);
            Properties.Settings.Default["ThirdNumSetting"] = Convert.ToDecimal(3);
            Properties.Settings.Default["FourthNumSetting"] = Convert.ToDecimal(8);

            DisplayCurrentSettings();
        }

        public ColumnsSettingsForm_ListS()
        {
            this.TopLevel = false;
            InitializeComponent();
            DisplayCurrentSettings();
        }
    }
}
