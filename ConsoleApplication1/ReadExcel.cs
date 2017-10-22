using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace ConsoleApplication1
{
    class ReadExcel
    {
        private List<List<string>> testSummaryInputs;
        private string[] qcFldr;
        private Dictionary<string, string> configDetails;
        public string inputFile;

        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet configSheet;
        Excel.Worksheet testSummarySheet;
        Excel.Worksheet bugSummarySheet;
        Excel.Range range;

        /*a property to get inputs for test, bug and config details.
         * data is read by the functions 'readTestInputsFromExcel'/'readBugInputsFromExcel/ readTestConfig' and
         * assigned to variables when the constructor of this class is called
        */
        public List<List<string>> TestInput
        {
            get { return testSummaryInputs; }
        }

        public string[] qcFldrNames
        {
            get { return qcFldr; }
        }

        public Dictionary<string, string> ConfigData
        {
            get { return configDetails; }
        }

        public ReadExcel(string inputFile)
        {
            /*A constructure that open the excel file and reads each sheet by calling functions.
             * It finally calls closeExcel() function to close the excel file.
             */
            this.inputFile = inputFile;
            xlApp = new Excel.Application();
            try
            {
                xlWorkBook = xlApp.Workbooks.Open(@inputFile, 0, true, 5, "", "", true,
                    Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            }
            catch(System.Runtime.InteropServices.COMException e)
            {
                Console.WriteLine("Invalid file name or location");
                Environment.Exit(0);
            }


            this.testSummaryInputs = readTestInputsFromExcel();
            this.configDetails = readTestConfig();
            this.qcFldr = readBugInputsFromExcel().Split(';');

            closeExcel();
        }


        private string readBugInputsFromExcel()
        {
            /*Reads the cell(1,2) of third sheet of excel file
             */
            bugSummarySheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);
            string qcFldrNames = (string)(bugSummarySheet.Cells[1, 2] as Excel.Range).Value;

            return qcFldrNames;

        }


        private List<List<string>> readTestInputsFromExcel()
        {
            /*reads 2nd sheet of the excel file. it return a list of lists. Each list contains release name, project name, test cycle id's,
             * status of the release, status color code and comments
             */
            int rCnt;

            testSummarySheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);

            range = testSummarySheet.UsedRange;
            int rw = range.Rows.Count;
            List<List<string>> testInputData = new List<List<string>>();

            /*row count starts from 2 to exclude the first row in 'Test Summary Inputs' sheet
             * This adds Release, Project, test cycle id's, status of the release, color to represent the status and
             * comments to the list
            */
            for (rCnt = 2; rCnt <= rw; rCnt++)
            {
                List<string> lst = new List<string>();
                //Release
                lst.Add((string)(range.Cells[rCnt, 1] as Excel.Range).Text);
                //Project
                lst.Add((string)(range.Cells[rCnt, 2] as Excel.Range).Text);
                //Cycle ids
                lst.Add((string)(range.Cells[rCnt, 3] as Excel.Range).Text);
                //Status
                lst.Add((string)(range.Cells[rCnt, 4] as Excel.Range).Text);
                //Color code
                lst.Add((string)(range.Cells[rCnt, 5] as Excel.Range).Text);
                //Comments
                lst.Add((string)(range.Cells[rCnt, 6] as Excel.Range).Text);

                testInputData.Add(lst);

            }

            return testInputData;
        }

        public Dictionary<string, string> readTestConfig()
        {
            /*Reads the first sheet to collect qc and other test settings
             * it reads QC url, domain, user, password, location to save reports, test status values (Eg: passed, no run..)
             * test priority values (Eg: Essential, High)
             */
            Dictionary<string, string> testConfig = new Dictionary<string, string>();

            int rCnt;
            configSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = configSheet.UsedRange;
            int rw = range.Rows.Count;
            
            for (rCnt = 1; rCnt <= rw; rCnt++)
            {

                testConfig.Add((string)(range.Cells[rCnt, 1] as Excel.Range).Value
                , (string)(range.Cells[rCnt, 2] as Excel.Range).Value);

            }

            return testConfig;

        }

        public void closeExcel()
        {
            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(configSheet);
            Marshal.ReleaseComObject(testSummarySheet);
            Marshal.ReleaseComObject(bugSummarySheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

    }
}
