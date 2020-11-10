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
            titleFontStyleBox.SelectedIndex = titleFontStyleBox.Items.IndexOf(Properties.Settings.Default.TitleFontTypeSetting);


            //Subtitle Settings Display Setup
            subtitleFontNum.Value = Properties.Settings.Default.SubtitleFontSetting; // Default Font Set to 10;
            subtitleBoldCheckBox.Checked = Properties.Settings.Default.SubtitleBoldSetting; // Default Bold True
            subtitleItalicCheckBox.Checked = Properties.Settings.Default.SubtitleItalicSetting; // Default Italic False
            subtitleUnderlineCheckBox.Checked = Properties.Settings.Default.SubtitleUnderlineSetting; // Default Underline False
            subtitleFontStyleBox.SelectedIndex = subtitleFontStyleBox.Items.IndexOf(Properties.Settings.Default.SubtitleFontTypeSetting);

            //Table Settings Display Setup
            tableFontNum.Value = Properties.Settings.Default.TableFontSetting; // Default Font Set to 11;
            tableBoldCheckBox.Checked = Properties.Settings.Default.TableBoldSetting; // Default Bold False
            tableItalicCheckBox.Checked = Properties.Settings.Default.TableItalicSetting; // Default Italic False
            tableUnderlineCheckBox.Checked = Properties.Settings.Default.TableUnderlineSetting; // Default Underline False
            tableFontStyleBox.SelectedIndex = tableFontStyleBox.Items.IndexOf(Properties.Settings.Default.TableFontTypeSetting);
        }

        public void SaveSettings ()
        {
            //Update Title Settings
            Properties.Settings.Default["TitleFontSetting"] = titleFontNum.Value;
            Properties.Settings.Default["TitleBoldSetting"] = titleBoldCheckBox.Checked;
            Properties.Settings.Default["TitleItalicSetting"] = titleItalicCheckBox.Checked;
            Properties.Settings.Default["TitleUnderlineSetting"] = titleUnderlineCheckBox.Checked;
            Properties.Settings.Default["TitleFontTypeSetting"] = titleFontStyleBox.SelectedItem;
            //Console.WriteLine(Properties.Settings.Default["TitleFontTypeSetting"]);

            //Subtitle Settings Display Setup
            Properties.Settings.Default["SubtitleFontSetting"] = subtitleFontNum.Value;
            Properties.Settings.Default["SubtitleBoldSetting"] = subtitleBoldCheckBox.Checked;
            Properties.Settings.Default["SubtitleItalicSetting"] = subtitleItalicCheckBox.Checked;
            Properties.Settings.Default["SubtitleUnderlineSetting"] = subtitleUnderlineCheckBox.Checked;
            Properties.Settings.Default["SubtitleFontTypeSetting"] = subtitleFontStyleBox.SelectedItem;

            //Table Settings Display Setup
            Properties.Settings.Default["TableFontSetting"] = tableFontNum.Value;
            Properties.Settings.Default["TableBoldSetting"] = tableBoldCheckBox.Checked;
            Properties.Settings.Default["TableItalicSetting"] = tableItalicCheckBox.Checked;
            Properties.Settings.Default["TableUnderlineSetting"] = tableUnderlineCheckBox.Checked;
            Properties.Settings.Default["TableFontTypeSetting"] = tableFontStyleBox.SelectedItem;
        }

        public void RestoreDefaults()
        {
            //Update Title Settings
            Properties.Settings.Default["TitleFontSetting"] = Convert.ToDecimal(14);
            Properties.Settings.Default["TitleBoldSetting"] = true;
            Properties.Settings.Default["TitleItalicSetting"] = false;
            Properties.Settings.Default["TitleUnderlineSetting"] = false;

            //Subtitle Settings Display Setup
            Properties.Settings.Default["SubtitleFontSetting"] = Convert.ToDecimal(10);
            Properties.Settings.Default["SubtitleBoldSetting"] = true;
            Properties.Settings.Default["SubtitleItalicSetting"] = false;
            Properties.Settings.Default["SubtitleUnderlineSetting"] = false;

            //Table Settings Display Setup
            Properties.Settings.Default["TableFontSetting"] = Convert.ToDecimal(11);
            Properties.Settings.Default["TableBoldSetting"] = false;
            Properties.Settings.Default["TableItalicSetting"] = false;
            Properties.Settings.Default["TableUnderlineSetting"] = false;

            DisplayCurrentSettings();
        }

        public TextSettingsForm_ListS()
        {
            this.TopLevel = false;
            InitializeComponent();

            fillBoxes();
            DisplayCurrentSettings();
            


        }

        public void fillBoxes()
        {
            var theFont = new Font("Arial", 12);
            if (titleFontStyleBox.Items.Count > 0)
            {
                titleFontStyleBox.Items.Clear();
            }

            if (subtitleFontStyleBox.Items.Count > 0)
            {
                subtitleFontStyleBox.Items.Clear();
            }

            if (tableFontStyleBox.Items.Count > 0)
            {
                tableFontStyleBox.Items.Clear();
            }

            foreach (var oneFontFamily in FontFamily.Families)
            {
                if (oneFontFamily.IsStyleAvailable(FontStyle.Regular))
                {
                    theFont = new Font(oneFontFamily.Name, 12);
                }
                else if (oneFontFamily.IsStyleAvailable(FontStyle.Bold))
                {
                    theFont = new Font(oneFontFamily.Name, 12, FontStyle.Bold);
                }
                else if (oneFontFamily.IsStyleAvailable(FontStyle.Italic))
                {
                    theFont = new Font(oneFontFamily.Name, 12, FontStyle.Italic);
                }
                else if (oneFontFamily.IsStyleAvailable(FontStyle.Strikeout))
                {
                    theFont = new Font(oneFontFamily.Name, 12, FontStyle.Strikeout);
                }
                else if (oneFontFamily.IsStyleAvailable(FontStyle.Underline))
                {
                    theFont = new Font(oneFontFamily.Name, 12, FontStyle.Underline);
                }


                if (theFont != null)
                {
                    titleFontStyleBox.Items.Add(theFont);
                    subtitleFontStyleBox.Items.Add(theFont);
                    tableFontStyleBox.Items.Add(theFont);

                }

            }
        }

        private void titleFontStyleBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            var theObject = new Font("Arial", 12);

            if (e.Index==-1)
            {
                return;
            }

            e.DrawBackground();

            if (e.State != DrawItemState.Checked && DrawItemState.Focus != DrawItemState.Checked)
            {
                e.DrawFocusRectangle();
            }

            var b = new SolidBrush(e.ForeColor);
            theObject = (titleFontStyleBox.Items[e.Index] as Font);

            StringFormat f1 = new StringFormat();
            f1.Trimming = StringTrimming.EllipsisWord;
            e.Graphics.DrawString(theObject.Name, theObject, b,e.Bounds, f1);

        }

        private void subtitleFontStyleBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            var theObject = new Font("Arial", 12);

            if (e.Index == -1)
            {
                return;
            }

            e.DrawBackground();

            if (e.State != DrawItemState.Checked && DrawItemState.Focus != DrawItemState.Checked)
            {
                e.DrawFocusRectangle();
            }

            var b = new SolidBrush(e.ForeColor);
            theObject = (subtitleFontStyleBox.Items[e.Index] as Font);

            StringFormat f1 = new StringFormat();
            f1.Trimming = StringTrimming.EllipsisWord;
            e.Graphics.DrawString(theObject.Name, theObject, b, e.Bounds, f1);
        }

        private void tableFontStyleBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            var theObject = new Font("Arial", 12);

            if (e.Index == -1)
            {
                return;
            }

            e.DrawBackground();

            if (e.State != DrawItemState.Checked && DrawItemState.Focus != DrawItemState.Checked)
            {
                e.DrawFocusRectangle();
            }

            var b = new SolidBrush(e.ForeColor);
            theObject = (tableFontStyleBox.Items[e.Index] as Font);

            StringFormat f1 = new StringFormat();
            f1.Trimming = StringTrimming.EllipsisWord;
            e.Graphics.DrawString(theObject.Name, theObject, b, e.Bounds, f1);
        }
    }
}
