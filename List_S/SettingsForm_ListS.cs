using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Svg;

namespace ListS
{
    public partial class SettingsForm_ListS : Form
    {
        //Setup Application movement
        private bool mouseDown;
        private Point lastLocation;
        private TextSettingsForm_ListS textSettings = new TextSettingsForm_ListS();
        private ColumnsSettingsForm_ListS colSettings = new ColumnsSettingsForm_ListS();
        private MiscSettingsForm_ListS miscSettings = new MiscSettingsForm_ListS();

        private void FormMove_MouseDown(MouseEventArgs e)
        {
            mouseDown = true;
            lastLocation = e.Location;
        }

        private void FormMove_MouseMove(MouseEventArgs e)
        {
            if (mouseDown)
            {
                this.Location = new Point((this.Location.X - lastLocation.X) + e.X, (this.Location.Y - lastLocation.Y) + e.Y);
                this.Update();
            }
        }

        private void FormMove_MouseUp()
        {
            mouseDown = false;
        }

        private void SettingsForm_ListS_MouseDown(object sender, MouseEventArgs e)
        {
            FormMove_MouseDown(e);
        }

        private void SettingsForm_ListS_MouseMove(object sender, MouseEventArgs e)
        {
            FormMove_MouseMove(e);
        }

        private void SettingsForm_ListS_MouseUp(object sender, MouseEventArgs e)
        {
            FormMove_MouseUp();
        }

        private void label1_MouseDown(object sender, MouseEventArgs e)
        {
            FormMove_MouseDown(e);
        }

        private void label1_MouseMove(object sender, MouseEventArgs e)
        {
            FormMove_MouseMove(e);
        }

        private void label1_MouseUp(object sender, MouseEventArgs e)
        {
            FormMove_MouseUp();
        }

        private void pictureBox1_MouseDown(object sender, MouseEventArgs e)
        {
            FormMove_MouseDown(e);
        }

        private void pictureBox1_MouseMove(object sender, MouseEventArgs e)
        {
            FormMove_MouseMove(e);
        }

        private void pictureBox1_MouseUp(object sender, MouseEventArgs e)
        {
            FormMove_MouseUp();
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            FormMove_MouseDown(e);
        }

        private void panel1_MouseMove(object sender, MouseEventArgs e)
        {
            FormMove_MouseMove(e);
        }

        private void panel1_MouseUp(object sender, MouseEventArgs e)
        {
            FormMove_MouseUp();
        }

        // Create a Drop Shadow for the Form
        private const int CS_DropShadow = 0x00020000;

        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ClassStyle = CS_DropShadow;
                return cp;
            }

        }

        /* 

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

        }*/

        public SettingsForm_ListS()
        {
            InitializeComponent();

            SvgDocument doc = SvgDocument.Open(@"C:\Users\Mr. Pickwick\Desktop\ListS_Logo36x36_noShadow.svg");
            Bitmap bmp = doc.Draw();
            pictureBox1.Image = bmp;

            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.closeButton.Image = (Image)(new Bitmap(this.closeButton.Image, new Size(16, 16)));

            
            splitContainer2.Panel1.Controls.Add(textSettings);
            splitContainer2.Panel1.Controls.Add(colSettings);
            splitContainer2.Panel1.Controls.Add(miscSettings);
            textSettings.Show();
            
        }

        private void closeButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        //Settings Sub-Menus Selections
        private void textSettingsButton_Click(object sender, EventArgs e)
        {
            colSettings.Hide();
            miscSettings.Hide();
            textSettings.Show();
        }

        private void columnsSettingsButton_Click(object sender, EventArgs e)
        {
            textSettings.Hide();
            miscSettings.Hide();
            colSettings.Show();
        }

        private void miscSettingsButton_Click(object sender, EventArgs e)
        {
            textSettings.Hide();
            colSettings.Hide();
            miscSettings.Show();
        }

        private void saveButton_Click(object sender, EventArgs e)
        {
            textSettings.SaveSettings();
            colSettings.SaveSettings();

            Properties.Settings.Default.Save();

            Form1.UpdateSettings();

            this.Close();
        }

        private void cancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
