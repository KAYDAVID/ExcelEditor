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
    public partial class TempSettingMenu : Form
    {
        public TempSettingMenu()
        {
            InitializeComponent();
            DisplayCurrentSettings();

        }

        // Display the current settings values.
        public void DisplayCurrentSettings ()
        {
            //Title Settings Display Setup
            titleFontNum.Value = Properties.Settings.Default.TitleFontSetting; // Default Font Set to 14;
            titleBoldCheckBox.Checked = Properties.Settings.Default.TitleBoldSetting; // Default Bold True
            titleItalicCheckBox.Checked = Properties.Settings.Default.TitleItalicSetting; // Default Italic False
            titleUnderlineCheckBox.Checked = Properties.Settings.Default.TitleUnderlineSetting; // Default Underline False

            //Table Settings Display Setup
            tableFontNum.Value = Properties.Settings.Default.TableFontSetting; // Default Font Set to 14;
            tableBoldCheckBox.Checked = Properties.Settings.Default.TableBoldSetting; // Default Bold True
            tableItalicCheckBox.Checked = Properties.Settings.Default.TableItalicSetting; // Default Italic False
            tableUnderlineCheckBox.Checked = Properties.Settings.Default.TableUnderlineSetting; // Default Underline False

            //Column Name Display Setup
            firstNameTextBox.Text = Properties.Settings.Default.FirstNameSetting; // Default ATTENDED 0
            secondNameTextBox.Text = Properties.Settings.Default.SecondNameSetting; // Default STATUS 8
            thirdNameTextBox.Text = Properties.Settings.Default.ThirdNameSetting; // Default FIRST NAME 2
            fourthNameTextBox.Text = Properties.Settings.Default.FourthNameSetting; // Default LAST NAME 3

            //Column Number Display Setup
            firstNumSelect.Value = Properties.Settings.Default.FirstNumSetting; // Default Column 99          
            secondNumSelect.Value = Properties.Settings.Default.SecondNumSetting; // Default Column 8 
            thirdNumSelect.Value = Properties.Settings.Default.ThirdNumSetting; // Default Column 2
            fourthNumSelect.Value = Properties.Settings.Default.FourthNumSetting; // Default Column 3

            //Number of empty rows Display Setup
            numERowsSelect.Value = Properties.Settings.Default.NumEmptySetting; // Default Empty row num 5 

        }

        // Save Button
        private void button1_Click(object sender, EventArgs e)
        {
            //Update Title Settings
            Properties.Settings.Default["TitleFontSetting"] = titleFontNum.Value;
            Properties.Settings.Default["TitleBoldSetting"] = titleBoldCheckBox.Checked;
            Properties.Settings.Default["TitleItalicSetting"] = titleItalicCheckBox.Checked;
            Properties.Settings.Default["TitleUnderlineSetting"] = titleUnderlineCheckBox.Checked;

            //Table Settings Display Setup
            Properties.Settings.Default["TableFontSetting"] = tableFontNum.Value;
            Properties.Settings.Default["TableBoldSetting"] = tableBoldCheckBox.Checked;
            Properties.Settings.Default["TableItalicSetting"] = tableItalicCheckBox.Checked;
            Properties.Settings.Default["TableUnderlineSetting"] = tableUnderlineCheckBox.Checked;

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

            //Number of empty rows Display Setup
            Properties.Settings.Default["NumEmptySetting"] = numERowsSelect.Value;

            Properties.Settings.Default.Save();

            Form1.UpdateSettings();
            

            this.Close();

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void tableFontNum_ValueChanged(object sender, EventArgs e)
        {

        }
    }
}
