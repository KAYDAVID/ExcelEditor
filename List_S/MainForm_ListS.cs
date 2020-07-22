using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.OleDb;
using Svg;

namespace ListS
{
    public partial class Form1 : Form
    {
        public List<string> listStatus = new List<string>();
        public List<string> firstName = new List<string>();
        public List<string> lastName = new List<string>();
        public bool fileIsRead = false;

        // Setup local variables for Settings
        //Title Settings
        private static int titleFont = (int)Properties.Settings.Default.TitleFontSetting; // Default Font Set to 14;
        private static bool titleBold = Properties.Settings.Default.TitleBoldSetting; // Default Bold True
        private static bool titleItalic = Properties.Settings.Default.TitleItalicSetting; // Default Italic False
        private static bool titleUnderline = Properties.Settings.Default.TitleUnderlineSetting; // Default Underline False

        //Subtitle Settings
        private static int subtitleFont = (int)Properties.Settings.Default.SubtitleFontSetting; // Default Font Set to 10;
        private static bool subtitleBold = Properties.Settings.Default.SubtitleBoldSetting; // Default Bold True
        private static bool subtitleItalic = Properties.Settings.Default.SubtitleItalicSetting; // Default Italic False
        private static bool subtitleUnderline = Properties.Settings.Default.SubtitleUnderlineSetting; // Default Underline False

        //Table Settings
        private static int tableFont = (int)Properties.Settings.Default.TableFontSetting; // Default Font Set to 11;
        private static bool tableBold = Properties.Settings.Default.TableBoldSetting; // Default Bold True
        private static bool tableItalic = Properties.Settings.Default.TableItalicSetting; // Default Italic False
        private static bool tableUnderline = Properties.Settings.Default.TableUnderlineSetting; // Default Underline False

        //Column Name
        private static string firstColumnName  = Properties.Settings.Default.FirstNameSetting; // Default ATTENDED 0
        private static string secondColumnName = Properties.Settings.Default.SecondNameSetting; // Default STATUS 8
        private static string thirdColumnName  = Properties.Settings.Default.ThirdNameSetting; // Default FIRST NAME 2
        private static string fourthColumnName = Properties.Settings.Default.FourthNameSetting; // Default LAST NAME 3

        //Column Number : switch to count from 0
        private static int firstColumnNum  = (int)Properties.Settings.Default.FirstNumSetting - 1; // Default Column 99          
        private static int secondColumnNum = (int)Properties.Settings.Default.SecondNumSetting - 1; // Default Column 8 
        private static int thirdColumnNum  = (int)Properties.Settings.Default.ThirdNumSetting - 1; // Default Column 2
        private static int fourthColumnNum = (int)Properties.Settings.Default.FourthNumSetting - 1; // Default Column 3

        //Number of empty rows
        private static int numOfERows = (int)Properties.Settings.Default.NumEmptySetting - 1; //switch to count from 0

        //Setup Application movement
        private bool mouseDown;
        private Point lastLocation;

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

        private void Form1_MouseDown(object sender, MouseEventArgs e)
        {
            FormMove_MouseDown(e);
        }

        private void Form1_MouseMove(object sender, MouseEventArgs e)
        {
            FormMove_MouseMove(e);
        }

        private void Form1_MouseUp(object sender, MouseEventArgs e)
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

        public Form1()
        {
            InitializeComponent();

            SvgDocument doc = SvgDocument.Open(@"C:\Users\Mr. Pickwick\Desktop\ListS_Logo36x36_noShadow.svg");
            Bitmap bmp = doc.Draw();
            pictureBox1.Image = bmp;

            //this.pictureBox1.Image = (Image)(new Bitmap(this.pictureBox1.Image, new Size(36, 36)));
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.closeButton.Image = (Image)(new Bitmap(this.closeButton.Image, new Size(16, 16)));
            this.settingsButton.Image = (Image)(new Bitmap(this.settingsButton.Image, new Size(16, 16)));

            pathTextBox.ReadOnly = true;
            //dateTimeTextBox.Text = DateTime.Now.ToString("dddd MMMM d, yyyy @ h"); //Saturday September 8, 2018 @ 10:00 AM

            //dateTimeTextBox.Text = dateTimeTextBox.Text + DateTime.Now.ToString("tt ");

            
        } 

        // Update all the local settings variables after altering them in the Settings Menu.
        public static void UpdateSettings()
        {
            //Title Settings
            titleFont = (int)Properties.Settings.Default.TitleFontSetting; // Default Font Set to 14;
            titleBold = Properties.Settings.Default.TitleBoldSetting; // Default Bold True
            titleItalic = Properties.Settings.Default.TitleItalicSetting; // Default Italic False
            titleUnderline = Properties.Settings.Default.TitleUnderlineSetting; // Default Underline False

            //Subtitle Settings
            subtitleFont = (int)Properties.Settings.Default.SubtitleFontSetting; // Default Font Set to 10;
            subtitleBold = Properties.Settings.Default.SubtitleBoldSetting; // Default Bold True
            subtitleItalic = Properties.Settings.Default.SubtitleItalicSetting; // Default Italic False
            subtitleUnderline = Properties.Settings.Default.SubtitleUnderlineSetting; // Default Underline False

            //Table Settings
            tableFont = (int)Properties.Settings.Default.TableFontSetting; // Default Font Set to 14;
            tableBold = Properties.Settings.Default.TableBoldSetting; // Default Bold True
            tableItalic = Properties.Settings.Default.TableItalicSetting; // Default Italic False
            tableUnderline = Properties.Settings.Default.TableUnderlineSetting; // Default Underline False

            //Column Name
            firstColumnName = Properties.Settings.Default.FirstNameSetting; // Default ATTENDED 0
            secondColumnName = Properties.Settings.Default.SecondNameSetting; // Default STATUS 8
            thirdColumnName = Properties.Settings.Default.ThirdNameSetting; // Default FIRST NAME 2
            fourthColumnName = Properties.Settings.Default.FourthNameSetting; // Default LAST NAME 3

            //Column Number : switch to count from 0
            firstColumnNum = (int)Properties.Settings.Default.FirstNumSetting - 1; // Default Column 99          
            secondColumnNum = (int)Properties.Settings.Default.SecondNumSetting - 1; // Default Column 8 
            thirdColumnNum = (int)Properties.Settings.Default.ThirdNumSetting - 1; // Default Column 2
            fourthColumnNum = (int)Properties.Settings.Default.FourthNumSetting - 1; // Default Column 3

            //Number of empty rows
            numOfERows = (int)Properties.Settings.Default.NumEmptySetting - 1; //switch to count from 0
        }

        private void SetTitleDimensions(Excel.Range myMergedCells)
        {
            // autofit does not set the row height if the row contains merged cells.
            // this is adapted from a post by Peer Moretti on MSDN forums.

            Excel.Range mySingleCell = myMergedCells.get_Offset(0, 100);
            double dblPointsWidth;

            // set our variables to the ranges concerned
            // and make sure the single cell has exactly the same font style, size, etc.
            // work out the width of the merged cells in points
            // this requires myMergedCells to be all the columns

            dblPointsWidth = 0;
            mySingleCell.Font.Name = myMergedCells.Font.Name;
            mySingleCell.Font.Size = myMergedCells.Font.Size;
            mySingleCell.Font.Bold = myMergedCells.Font.Bold;
            
            foreach (Excel.Range myColumn in myMergedCells.Columns)
            {
                dblPointsWidth = dblPointsWidth + (double)myColumn.Width;
            }
            
            // set the single cell to the same width as the merged cells in points
            mySingleCell.ColumnWidth = (double)mySingleCell.ColumnWidth * dblPointsWidth / (double)mySingleCell.Width;
            double tooSmall = mySingleCell.ColumnWidth;

            // put the information from the merged cell in to the single cell
            mySingleCell.Value2 = myMergedCells.Value2;

            // now autofit
            mySingleCell.Rows.AutoFit();
            mySingleCell.Columns.AutoFit();

            if (mySingleCell.ColumnWidth < tooSmall/2)
            {
                myMergedCells.RowHeight = mySingleCell.RowHeight / 2;
                myMergedCells.ColumnWidth = tooSmall/(myMergedCells.Columns.Count -1);

                mySingleCell.Columns.Hidden = true;
                return;
            }

            // set the height of the merged cell to the height of the single cell
            myMergedCells.RowHeight = mySingleCell.RowHeight/2.7;
            myMergedCells.ColumnWidth = mySingleCell.ColumnWidth/2.8;

            mySingleCell.Columns.Hidden = true;
        }

        //Get the Class name and Date from text boxes and format it.
        //If the Name and Date text boxes are empty insert placeholder text.
        private string CreateTitleString()
        {
            //Type class name on the first line
            string firstLine = "ClassName";
            firstLine = classNameTextBox.Text;
            if (classNameTextBox.Text == " Class Name" || classNameTextBox.Text == string.Empty)
            {
                firstLine = "Default Class Heading: Default Class Sub-Heading";
            }

            //Tap Alt+Enter to create a second line 
            string secondLine = "Date";
            secondLine = dateTimeTextBox.Text;
            if (dateTimeTextBox.Text == " Date and Time" || classNameTextBox.Text == string.Empty)
            {
                secondLine = "Defaultday February 30th, 1895 @ 10:00 AM";
            }

            string allLine = firstLine + "\n" + secondLine;
            return allLine;
        }

        //Create Button
        private void createButton_Click(object sender, EventArgs e)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            if (fileIsRead == false)
            {
                MessageBox.Show("Please choose a file.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            this.Cursor = Cursors.WaitCursor;
            Application.UseWaitCursor = true;
            Application.DoEvents();

            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = false;

                //Get a new Workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                string endColumn = "D" + "2"; // Alterable for testing purposes

                //Select these two rows A-D and click Merge and Center
                oSheet.get_Range("A1", endColumn).MergeCells = true;
                oSheet.get_Range("A1", endColumn).VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("A1", endColumn).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;

                //Set title text settings : Default(Increase Font size to 14, Bold, : middle align/center done above).
                oSheet.get_Range("A1", endColumn).Font.Size = titleFont;
                oSheet.get_Range("A1", endColumn).Font.Bold = titleBold;
                oSheet.get_Range("A1", endColumn).Font.Italic = titleItalic;
                oSheet.get_Range("A1", endColumn).Font.Underline = titleUnderline;
                oSheet.get_Range("A1", endColumn).WrapText = true;

                //Get the Class name and Date from text boxes and format it.
                string allLine = CreateTitleString();

                //Take the compiled title in allLine put that in the sheet
                oSheet.get_Range("A1", endColumn).Value = allLine;

                //Alter the Cell to fit the data
                oRng = oSheet.get_Range("A1", endColumn);
                SetTitleDimensions(oRng);

                //Setup sub heading format
                //A3 : First Column
                oSheet.get_Range("A3", "A4").MergeCells = true;
                oSheet.get_Range("A3", "A4").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("A3", "A4").HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("A3").Value = firstColumnName;
                oSheet.get_Range("A3", "A4").Font.Size = subtitleFont;
                oSheet.get_Range("A3", "A4").Font.Bold = subtitleBold;
                oSheet.get_Range("A3", "A4").Font.Italic = subtitleItalic;
                oSheet.get_Range("A3", "A4").Font.Underline = subtitleUnderline;


                //B3 : Second Column
                oSheet.get_Range("B3", "B4").MergeCells = true;
                oSheet.get_Range("B3", "B4").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("B3", "B4").HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("B3").Value = secondColumnName;
                oSheet.get_Range("B3", "B4").Font.Size = subtitleFont;
                oSheet.get_Range("B3", "B4").Font.Bold = subtitleBold;
                oSheet.get_Range("B3", "B4").Font.Italic = subtitleItalic;
                oSheet.get_Range("B3", "B4").Font.Underline = subtitleUnderline;


                //C3 : Third Column
                oSheet.get_Range("C3", "C4").MergeCells = true;
                oSheet.get_Range("C3", "C4").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("C3", "C4").HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("C3").Value = thirdColumnName;
                oSheet.get_Range("C3", "C4").Font.Size = subtitleFont;
                oSheet.get_Range("C3", "C4").Font.Bold = subtitleBold;
                oSheet.get_Range("C3", "C4").Font.Italic = subtitleItalic;
                oSheet.get_Range("C3", "C4").Font.Underline = subtitleUnderline;


                //D3 : Fourth Column
                oSheet.get_Range("D3", "D4").MergeCells = true;
                oSheet.get_Range("D3", "D4").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("D3", "D4").HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("D3").Value = fourthColumnName;
                oSheet.get_Range("D3", "D4").Font.Size = subtitleFont;
                oSheet.get_Range("D3", "D4").Font.Bold = subtitleBold;
                oSheet.get_Range("D3", "D4").Font.Italic = subtitleItalic;
                oSheet.get_Range("D3", "D4").Font.Underline = subtitleUnderline;


                //Take what is in the lists and put in excel

                

                int buffer = 5;

                for (int i = 1; i < listStatus.Count(); i++)
                {
                    oSheet.get_Range("B" + buffer).Value = listStatus[i];
                    buffer++;
                }

                buffer = 5;
                for (int i = 1; i < firstName.Count(); i++)
                {
                    oSheet.get_Range("C" + buffer).Value = firstName[i];
                    buffer++;
                }

                buffer = 5;
                for (int i = 1; i < lastName.Count(); i++)
                {
                    oSheet.get_Range("D" + buffer).Value = lastName[i];
                    buffer++;
                }

                //Format data 
                //oSheet.get_Range(S, E).Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                //Highlight five rows below the last name on the list all borders
                buffer = 4; //to include the header which comes from the lists
                //Check for the longest list so no overwriting
                int listSize = listStatus.Count();
                if (listSize < firstName.Count())
                {
                    listSize = firstName.Count();
                }
                else if (listSize < lastName.Count())
                {
                    listSize = lastName.Count();
                }

                int end = buffer + listSize + numOfERows;
                string E = "D" + end.ToString();

                // Add all borders to the data 
                oSheet.get_Range("A3", E).Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                oSheet.get_Range("A5", E).Font.Size = tableFont;
                oSheet.get_Range("A5", E).Font.Bold = tableBold;
                oSheet.get_Range("A5", E).Font.Italic = tableItalic;
                oSheet.get_Range("A5", E).Font.Underline = tableUnderline;



                // Align all the data
                oSheet.get_Range("A3", E).HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;

                //Skip one row then in the next merge 14 bold align left
                int instruct = end + 2;
                string IS = "A" + instruct.ToString();
                string IE = "D" + instruct.ToString();

                oSheet.get_Range(IS, IE).MergeCells = true;
                oSheet.get_Range(IS, IE).Font.Bold = true;
                oSheet.get_Range(IS, IE).Font.Size = 14;

                oSheet.get_Range(IS, IE).Value = "Instructor:_____________________________";

                //Skip one row then in the next merge 14 bold align left
                int totalA = instruct + 2;
                string TAS = "A" + totalA.ToString();
                string TAE = "D" + totalA.ToString();

                oSheet.get_Range(TAS, TAE).MergeCells = true;
                oSheet.get_Range(TAS, TAE).Font.Bold = true;
                oSheet.get_Range(TAS, TAE).Font.Size = 14;

                oSheet.get_Range(TAS, TAE).Value = "Total Attendees:________________________";
                
                //Make sure Excel is visible and give user control of Microsoft Excel's Lifetime.
                oXL.Visible = true;
                oXL.UserControl = true;

            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
                //throw;
            }

            Application.UseWaitCursor = false;
            this.Cursor = Cursors.Default;

        }

        //Choose and Read button
        private void chooseFileButton_Click(object sender, EventArgs e)
        {
            string filePath = string.Empty;
            string fileExt = string.Empty;
            

            OpenFileDialog file = new OpenFileDialog();

            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filePath = file.FileName;
                fileExt = Path.GetExtension(filePath);
                if (fileExt.CompareTo(".csv") == 0)
                {
                    if (fileIsRead == true)
                    {
                        listStatus.Clear();
                        firstName.Clear();
                        lastName.Clear();
                    }
                        try
                        {
                            string[] data;
                            pathTextBox.Text = filePath; // Display path in textbox
                            data = System.IO.File.ReadAllLines(filePath); // Read All the lines of the file into the string array data

                            foreach (string L in data) // cut up the important data to lists
                            {
                                string[] cuttings = L.Split(',');
                                listStatus.Add(cuttings[secondColumnNum]);
                                firstName.Add(cuttings[thirdColumnNum]);
                                lastName.Add(cuttings[fourthColumnNum]);
                            }
                        fileIsRead = true;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message.ToString());
                            //throw;
                        }
                }

                else
                {
                    MessageBox.Show("Please choose a .csv file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }

        }

        //Settings Button
        private void settingsButton_Click(object sender, EventArgs e)
        {
            //TempSettingMenu F2 = new TempSettingMenu();
            SettingsForm_ListS F2 = new SettingsForm_ListS();
            F2.ShowDialog();
        }

        private void closeButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }


    }
}
