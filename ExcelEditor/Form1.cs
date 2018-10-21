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

namespace ExcelEditor
{
    public partial class Form1 : Form
    {
        public List<string> listStatus = new List<string>();
        public List<string> firstName = new List<string>();
        public List<string> lastName = new List<string>();
        public bool fileIsRead = false;

        public Form1()
        {
            InitializeComponent();
            pathTextBox.ReadOnly = true;
        }

        private void SetRowHeight(Excel.Range myMergedCells)
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

        private void button1_Click(object sender, EventArgs e)
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

                //Insert two new rows above Row 1
                //Select these two rows A-E and click Merge and Center
                oSheet.get_Range("A1", "D2").MergeCells = true;
                oSheet.get_Range("A1", "D2").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("A1", "D2").HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;

                //Increase Font size to 14, Bold, : middle align/center done above.
                oSheet.get_Range("A1", "D2").Font.Bold = true;
                oSheet.get_Range("A1", "D2").Font.Size = 14;
                oSheet.get_Range("A1", "D2").WrapText = true;
                //Type class name on the first line
                string firstLine = "ClassName";
                firstLine = classNameTextBox.Text;
                if (classNameTextBox.Text == string.Empty)
                {
                    firstLine = "Default Class Heading: Default Class Sub-Heading";
                }
                //Tap Alt+Enter to create a second line 
                string secondLine = "Date";
                secondLine = dateTimeTextBox.Text;
                if (dateTimeTextBox.Text == string.Empty)
                {
                    secondLine = "Saturday September 8, 2018 @ 10:00 AM";
                }

                string allLine = firstLine + "\n" + secondLine;
                oSheet.get_Range("A1", "D2").Value = allLine;

                oRng = oSheet.get_Range("A1", "D2");
                SetRowHeight(oRng);

                //Setup sub heading format
                //A3
                oSheet.get_Range("A3", "A4").MergeCells = true;
                oSheet.get_Range("A3", "A4").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("A3", "A4").HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("A3", "A4").Font.Bold = true;
                oSheet.get_Range("A3", "A4").Font.Size = 10;

                //B3
                oSheet.get_Range("B3", "B4").MergeCells = true;
                oSheet.get_Range("B3", "B4").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("B3", "B4").HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("B3", "B4").Font.Bold = true;
                oSheet.get_Range("B3", "B4").Font.Size = 10;

                //C3
                oSheet.get_Range("C3", "C4").MergeCells = true;
                oSheet.get_Range("C3", "C4").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("C3", "C4").HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("C3", "C4").Font.Bold = true;
                oSheet.get_Range("C3", "C4").Font.Size = 10;

                //D3
                oSheet.get_Range("D3", "D4").MergeCells = true;
                oSheet.get_Range("D3", "D4").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("D3", "D4").HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
                oSheet.get_Range("D3", "D4").Font.Bold = true;
                oSheet.get_Range("D3", "D4").Font.Size = 10;

                //Take what is in the lists and put in excel
                oSheet.get_Range("A3").Value = "ATTENDED";
                oSheet.get_Range("B3").Value = listStatus[0];
                oSheet.get_Range("C3").Value = firstName[0];
                oSheet.get_Range("D3").Value = lastName[0];

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

                int end = buffer + listSize + 4;
                string E = "D" + end.ToString();

                // Add all borders to the data 
                oSheet.get_Range("A3", E).Cells.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

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
        private void button2_Click(object sender, EventArgs e)
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
                            // Display path in textbox
                            pathTextBox.Text = filePath;
                            data = System.IO.File.ReadAllLines(filePath);

                            foreach (string L in data) // cut up the important data to lists
                            {
                                string[] cuttings = L.Split(',');
                                listStatus.Add(cuttings[0]);
                                firstName.Add(cuttings[1]);
                                lastName.Add(cuttings[2]);
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

    }
}
