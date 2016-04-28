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




namespace CourseScheduleVerifier2._0
{
    public partial class Form1 : Form
    {
        public struct courseDetails
        {


            public string instructor;
            public string room;
            public string day;
            public string time;

        }

        //keeps track of how many classes were added 
        int classCounter = 0;

        //Stores courseDetails structs in a list
        public List<courseDetails> courseList = new List<courseDetails>();

       
        char currColumn;
        int currRow;

        public Form1()
        {
            InitializeComponent();
        }

        private void Import_Click(object sender, EventArgs e)
        {
            //create instance of OpenFileDialog box
            OpenFileDialog ofd = new OpenFileDialog();

            //Filter to excel files only
            ofd.Filter = "Excel Files (.xlsx)|*.xlsx|All Files (*.*)|*.*";

            if (ofd.ShowDialog() == DialogResult.OK)
            {

                string filePath = ofd.FileName;
                
                //Loops through each column and row containing pertinent data to find conflicting classes
                for (int i = 4; i <= 13; i++)
                {
                    currRow = i;
                    if (GetCellValues(filePath, "Sheet1", ("B" + currRow).ToString()) == "CSC" || (GetCellValues(filePath, "Sheet1", ("B" + currRow).ToString()) == "SCI"))
                    {


                        courseDetails courseDetails1;


                        courseDetails1.instructor = ((GetCellValues(filePath, "Sheet1", "G" + currRow)) + (GetCellValues(filePath, "Sheet1", "H" + currRow)));
                        courseDetails1.room = ((GetCellValues(filePath, "Sheet1", "K" + currRow)) + (GetCellValues(filePath, "Sheet1", "L" + currRow)));
                        courseDetails1.day = ((GetCellValues(filePath, "Sheet1", "O" + currRow)) + (GetCellValues(filePath, "Sheet1", "P" + currRow)) + (GetCellValues(filePath, "Sheet1", "Q" + currRow)) + (GetCellValues(filePath, "Sheet1", "R" + currRow)) + (GetCellValues(filePath, "Sheet1", "S" + currRow)));
                        courseDetails1.time = ((GetCellValues(filePath, "Sheet1", "M" + currRow)) + (GetCellValues(filePath, "Sheet1", "N" + currRow)));

                        courseList.Add(courseDetails1);                       
                        
                    }
                }

                for (int j = 0; j < courseList.Count; j++)
                {
                    listBox1.Items.Add(courseList[j].instructor + courseList[j].room + courseList[j].day + courseList[j].time);

                }
            }
        }

        public static string GetCellValues(string fileName, string sheetName, string addressName)
        {
            string value = null;

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
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
                            // shared strings table.
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
            return value;

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
    }
}
