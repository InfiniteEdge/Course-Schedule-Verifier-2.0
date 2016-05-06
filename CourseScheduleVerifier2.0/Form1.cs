using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO.Packaging;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using Color = System.Drawing.Color;




namespace CourseScheduleVerifier2._0
{
    public partial class Form1 : Form
    {
        //Creates a struct to hold course details such as class name, room, instructor, etc
        public struct courseDetails
        {
            public string section;
            public string instructor;
            public string room;
            public string day;
            public string time;
            public string key;
            public string className;
            public Boolean tba;
        }

        //keeps track of how many classes were added 
        int classCounter = 0;

        //Stores courseDetails structs in a list
        public List<courseDetails> courseList = new List<courseDetails>();

       
        //Keeps track of the current row to grab all the data in the specified row
        int currRow;

        
        

        public Form1()
        {
            InitializeComponent();
        }

        private void Import_Click(object sender, EventArgs e)
        {
            remove.Enabled = true;
            clear.Enabled = true;
            courseList.Clear();
            classCounter = 0;
            

            //create instance of OpenFileDialog box
            OpenFileDialog ofd = new OpenFileDialog();

            //Filter to excel files only
            ofd.Filter = "Excel Files (.xlsx)|*.xlsx|All Files (*.*)|*.*";

            if (ofd.ShowDialog() == DialogResult.OK)
            {

                string filePath = ofd.FileName;
                
                if (bw.IsBusy != true)
                {
                    bw.RunWorkerAsync();
                }
               
                
                //Loops through each column and row containing pertinent data to find conflicting classes and adds those conflictions to a List Box
                for (int i = 4; i <= 17; i++)
                {
                    currRow = i;                  
                    if (GetCellValues(filePath, "Sheet1", ("B" + currRow).ToString()) == "CSC" || (GetCellValues(filePath, "Sheet1", ("B" + currRow).ToString()) == "SCI"))
                    {
                        
                        courseDetails courseDetails1;

                        courseDetails1.className = ((GetCellValues(filePath, "Sheet1", "B" + currRow)) + (GetCellValues(filePath, "Sheet1", "C" + currRow)));
                        courseDetails1.key = (classCounter + 1).ToString();
                        courseDetails1.section = ((GetCellValues(filePath, "Sheet1", "D" + currRow)));
                        courseDetails1.instructor = ((GetCellValues(filePath, "Sheet1", "G" + currRow)) + (GetCellValues(filePath, "Sheet1", "H" + currRow)));
                        courseDetails1.room = ((GetCellValues(filePath, "Sheet1", "K" + currRow)) + (GetCellValues(filePath, "Sheet1", "L" + currRow)));                      
                        if (((GetCellValues(filePath, "Sheet1", "K" + currRow)) == "tba") || ((GetCellValues(filePath, "Sheet1", "K" + currRow)) == "Tba") || ((GetCellValues(filePath, "Sheet1", "K" + currRow)) == "TBA"))
                        {
                            courseDetails1.tba = true;
                        }  
                        else
                        {
                            courseDetails1.tba = false;
                        }
                        courseDetails1.day = ((GetCellValues(filePath, "Sheet1", "O" + currRow)) + (GetCellValues(filePath, "Sheet1", "P" + currRow)) + (GetCellValues(filePath, "Sheet1", "Q" + currRow)) + (GetCellValues(filePath, "Sheet1", "R" + currRow)) + (GetCellValues(filePath, "Sheet1", "S" + currRow)));
                        courseDetails1.time = ((GetCellValues(filePath, "Sheet1", "M" + currRow)) + " -" + (GetCellValues(filePath, "Sheet1", "N" + currRow)));

                        if (courseDetails1.tba == false)
                        {
                            courseList.Add(courseDetails1);
                            classCounter = classCounter + 1;
                        }
                                                                    
                    }

                }
                
                Boolean conflicts = false;
                Boolean check;

                for (int i = 0; i < classCounter; i++)
                {
                    for (int j = 0 + i; j < classCounter; j++)
                    {
                        check = true;

                        if (i == j)
                        {
                            check = false;
                        }
                        if ((courseList[i].instructor == courseList[j].instructor) && (courseList[i].day == courseList[j].day) && (courseList[i].time == courseList[j].time) && (check == true))
                        {
                            listBox1.Items.Add(courseList[i].className + "-" + courseList[i].section + " instructor conflict with: " + courseList[j].className + "-" + courseList[j].section);
                            conflicts = true;
                        }
                        if ((courseList[i].room == courseList[j].room) && (courseList[i].day == courseList[j].day) && (courseList[i].time == courseList[j].time) && (check == true))
                        {
                            listBox1.Items.Add(courseList[i].className + "-" + courseList[i].section + " room conflict with: " + courseList[j].className + "-" + courseList[j].section);
                            conflicts = true;
                        }
                    }
                }

                if (conflicts == false)
                {
                    listBox1.Items.Add("There are no conflicting classes");
                }       
                                               
            }
            
        }

        public static string GetCellValues(string fileName, string sheetName, string addressName)
        {
            string value = null;
            using (Stream stream = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false))
                {
                    WorkbookPart wbPart = document.WorkbookPart;
                    Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                        Where(s => s.Name == sheetName).FirstOrDefault();

                    if (theSheet == null)
                    {
                        throw new ArgumentException("sheetName");
                    }

                    WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                    Cell theCell = wsPart.Worksheet.Descendants<Cell>().
                        Where(c => c.CellReference == addressName).FirstOrDefault();

                    if (theCell != null)
                    {
                        value = theCell.InnerText;
                    }

                    if (theCell.DataType != null)
                    {
                        switch (theCell.DataType.Value)
                        {
                            case CellValues.SharedString:

                                // For shared strings, look up the value in the
                                // shared strings table
                                var stringTable =
                                    wbPart.GetPartsOfType<SharedStringTablePart>()
                                    .FirstOrDefault();

                                // If the shared string table is missing, something 
                                // is wrong. Return the index that is in
                                // the cell. Otherwise, look up the correct text in 
                                // the table.
                                if (stringTable != null)
                                {
                                    value =
                                        stringTable.SharedStringTable
                                        .ElementAt(int.Parse(value)).InnerText;
                                }
                                break;

                            case CellValues.Boolean:
                                switch (value)
                                {
                                    case "0":
                                        value = "FALSE";
                                        break;
                                    default:
                                        value = "TRUE";
                                        break;
                                }
                                break;
                        }

                    }
                }
            }
            return value;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            fIleToolStripMenuItem.ForeColor = Color.Orange;
            clear.Enabled = false;
            remove.Enabled = false;          
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        //Clears the list box in order to remove all conflicting classes as they are fixed
        private void clear_Click(object sender, EventArgs e)
        {
            listBox1.Items.Clear();
        }

        //Removes the selected item in order to remove specific conflicting classes as they are fixed 
        private void remove_Click(object sender, EventArgs e)
        {
            int removeNumber = listBox1.SelectedIndex;
            listBox1.Items.Remove(listBox1.SelectedItem);
        }

        private void fIleToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void importToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Import_Click(sender, e);
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {
           
        }

        private void label2_Click(object sender, EventArgs e)
        {
            
        }

        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {          
            
        }     

        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            this.label2.Text = (e.ProgressPercentage.ToString() + "%");
            this.progressBar1.Increment(e.ProgressPercentage);
        }

        
    }
}
