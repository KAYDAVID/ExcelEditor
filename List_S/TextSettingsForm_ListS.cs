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
    public partial class TextSettingsForm_ListS : Form
    {

        public void DisplayCurrentSettings ()
        {
            //Title Settings Display Setup
            titleFontNum.Value = Properties.Settings.Default.TitleFontSetting; // Default Font Set to 14;
            titleBoldCheckBox.Checked = Properties.Settings.Default.TitleBoldSetting; // Default Bold True
            titleItalicCheckBox.Checked = Properties.Settings.Default.TitleItalicSetting; // Default Italic False
            titleUnderlineCheckBox.Checked = Properties.Settings.Default.TitleUnderlineSetting; // Default Underline False

            //Subtitle Settings Display Setup
            subtitleFontNum.Value = Properties.Settings.Default.SubtitleFontSetting; // Default Font Set to 10;
            subtitleBoldCheckBox.Checked = Properties.Settings.Default.SubtitleBoldSetting; // Default Bold True
            subtitleItalicCheckBox.Checked = Properties.Settings.Default.SubtitleItalicSetting; // Default Italic False
            subtitleUnderlineCheckBox.Checked = Properties.Settings.Default.SubtitleUnderlineSetting; // Default Underline False

            //Table Settings Display Setup
            tableFontNum.Value = Properties.Settings.Default.TableFontSetting; // Default Font Set to 11;
            tableBoldCheckBox.Checked = Properties.Settings.Default.TableBoldSetting; // Default Bold False
            tableItalicCheckBox.Checked = Properties.Settings.Default.TableItalicSetting; // Default Italic False
            tableUnderlineCheckBox.Checked = Properties.Settings.Default.TableUnderlineSetting; // Default Underline False
        }

        public void SaveSettings ()
        {
            //Update Title Settings
            Properties.Settings.Default["TitleFontSetting"] = titleFontNum.Value;
            Properties.Settings.Default["TitleBoldSetting"] = titleBoldCheckBox.Checked;
            Properties.Settings.Default["TitleItalicSetting"] = titleItalicCheckBox.Checked;
            Properties.Settings.Default["TitleUnderlineSetting"] = titleUnderlineCheckBox.Checked;

            //Subtitle Settings Display Setup
            Properties.Settings.Default["SubtitleFontSetting"] = subtitleFontNum.Value;
            Properties.Settings.Default["SubtitleBoldSetting"] = subtitleBoldCheckBox.Checked;
            Properties.Settings.Default["SubtitleItalicSetting"] = subtitleItalicCheckBox.Checked;
            Properties.Settings.Default["SubtitleUnderlineSetting"] = subtitleUnderlineCheckBox.Checked;

            //Table Settings Display Setup
            Properties.Settings.Default["TableFontSetting"] = tableFontNum.Value;
            Properties.Settings.Default["TableBoldSetting"] = tableBoldCheckBox.Checked;
            Properties.Settings.Default["TableItalicSetting"] = tableItalicCheckBox.Checked;
            Properties.Settings.Default["TableUnderlineSetting"] = tableUnderlineCheckBox.Checked;
        }

        public TextSettingsForm_ListS()
        {
            this.TopLevel = false;
            InitializeComponent();

            DisplayCurrentSettings();

        }

    }
}
