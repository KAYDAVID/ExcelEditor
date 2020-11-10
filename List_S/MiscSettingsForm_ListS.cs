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
    public partial class MiscSettingsForm_ListS : Form
    {
        public void DisplayCurrentSettings()
        {
            emptyRowsNumBox.Value = Properties.Settings.Default.NumEmptySetting; // Default 5

            // dateCheckBox.Checked = Properties.Settings.Default.
            // fileCheckBox.Checked = Properties.Settings.Default.

            //Signature Lines Display Setup
            firstLineTextBox.Text = Properties.Settings.Default.SecondNameSetting; // Default STATUS 8
            secondLineTextBox.Text = Properties.Settings.Default.FirstNameSetting; // Default ATTENDED 0

            ////Column Number Display Setup
            //firstNumSelect.Value = Properties.Settings.Default.FirstNumSetting; // Default Column 99          
            //secondNumSelect.Value = Properties.Settings.Default.SecondNumSetting; // Default Column 8 
            //thirdNumSelect.Value = Properties.Settings.Default.ThirdNumSetting; // Default Column 2
            //fourthNumSelect.Value = Properties.Settings.Default.FourthNumSetting; // Default Column 3
        }

        public void SaveSettings()
        {
            ////Column Name Display Setup
            //Properties.Settings.Default["FirstNameSetting"] = secondLineTextBox.Text;
            //Properties.Settings.Default["SecondNameSetting"] = secondNameTextBox.Text;
            //Properties.Settings.Default["ThirdNameSetting"] = thirdNameTextBox.Text;
            //Properties.Settings.Default["FourthNameSetting"] = fourthNameTextBox.Text;

            ////Column Number Display Setup
            //Properties.Settings.Default["FirstNumSetting"] = firstNumSelect.Value;
            //Properties.Settings.Default["SecondNumSetting"] = secondNumSelect.Value;
            //Properties.Settings.Default["ThirdNumSetting"] = thirdNumSelect.Value;
            //Properties.Settings.Default["FourthNumSetting"] = fourthNumSelect.Value;
        }

        public void RestoreDefaults()
        {
            ////Column Name Display Setup
            //Properties.Settings.Default["FirstNameSetting"] = secondLineTextBox.Text;
            //Properties.Settings.Default["SecondNameSetting"] = secondNameTextBox.Text;
            //Properties.Settings.Default["ThirdNameSetting"] = thirdNameTextBox.Text;
            //Properties.Settings.Default["FourthNameSetting"] = fourthNameTextBox.Text;

            ////Column Number Display Setup
            //Properties.Settings.Default["FirstNumSetting"] = firstNumSelect.Value;
            //Properties.Settings.Default["SecondNumSetting"] = secondNumSelect.Value;
            //Properties.Settings.Default["ThirdNumSetting"] = thirdNumSelect.Value;
            //Properties.Settings.Default["FourthNumSetting"] = fourthNumSelect.Value;
        }
        public MiscSettingsForm_ListS()
        {
            this.TopLevel = false;
            InitializeComponent();
        }
    }
}
