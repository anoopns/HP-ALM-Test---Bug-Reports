using System;
using System.Data;
using TDAPIOLELib;
using System.Collections.Generic;


namespace ConsoleApplication1
{
    class Program
    {

        static void Main(string[] args)
        {
            
            Console.WriteLine("Enter input file location? ");
            string input_file = Console.ReadLine();

            Console.WriteLine("Do you want to generate ARQ report?(Y/N)");
            string response = Console.ReadLine();

            Boolean arqResp = false;

            if(response == "Y")
            {
                arqResp = true;
            }

            //Read excel for config, test and bug inputs 
            ReadExcel readExcel = new ReadExcel(input_file);

            //Read config details
            Dictionary<string, string> configDetails = readExcel.ConfigData;
            string qcUrl = configDetails["qcURL"];
            string qcDomain = configDetails["qcDomain"];
            string qcProject = configDetails["qcProject"];
            string qcUser = configDetails["qcUser"];
            string qcPassword = configDetails["qcPassword"];
            string reportLocation = configDetails["reportLocation"];
            string[] testStatusValues = configDetails["tcStatusValues"].Split(';');
            string[] tcPriorityValues = configDetails["tcPriorityValues"].Split(',');

            Connect con = new Connect(qcDomain, qcProject, qcUser, qcPassword, qcUrl);
            TDConnection tdc = con.connectToProject();

            //Read test data
            List<List<string>> releaseData = new List<List<string>>();
            releaseData = readExcel.TestInput;

            //Read ARQ data
            string[] qcFldrNames = readExcel.qcFldrNames;


            Dictionary<string, string> releaseColorCode = new Dictionary<string, string>();
            
            if(tdc.Connected)
            {
                
                TestData tData = new TestData(tdc);
                DataTable masterTestData = tData.createMasterTestTable(releaseData);
				DataSet testStatusDtSt = tData.createTestReportByStatus(releaseData, masterTestData, testStatusValues);
				DataSet testPriorityDtSt = tData.createTestReportBySeverity(releaseData, masterTestData, testStatusValues, tcPriorityValues);

				DataTable summaryTable = new DataTable("Test_Summary");
                summaryTable.Columns.Add("Release");
                summaryTable.Columns.Add("Project");
                summaryTable.Columns.Add("Total Test Cases", typeof(Int32));
                summaryTable.Columns.Add("Total % Completed", typeof(float));
                summaryTable.Columns.Add("% Cond. Passed & Passed", typeof(float));
                summaryTable.Columns.Add("Status", typeof(string));
                summaryTable.Columns.Add("Comments", typeof(string));
                
                List<string> releases = new List<string>();
                List<string> projects = new List<string>();
                List<string> releaseStatus = new List<string>();
                List<string> comments = new List<string>();

                foreach (List<string> item in releaseData)
                {
                    if (item[0] != null & item[1] != null )
                    {
                        releases.Add(item[0]);
                        projects.Add(item[1]);
                    }

                    if(item[3] != null)
                    {
                        releaseStatus.Add(item[3]);
                    }
                    else
                    {
                        releaseStatus.Add(" ");
                    }
                    if (item[5] != null)
                    {
                        comments.Add(item[5]);
                    }
                    else
                    {
                        comments.Add(" ");
                    }

                }


                for(int i = 0; i < releases.Count; i++)
                {

                    int totalTestCases = testStatusDtSt.Tables[i].Rows[testStatusDtSt.Tables[i].Rows.Count - 1].Field<Int32>("Total Tests");
                    int descopedTests = testStatusDtSt.Tables[i].Rows[testStatusDtSt.Tables[i].Rows.Count - 1].Field<Int32>("Descoped");
                    int carriedFwdTests = testStatusDtSt.Tables[i].Rows[testStatusDtSt.Tables[i].Rows.Count - 1].Field<Int32>("Carried Forward");
                    int total = totalTestCases - descopedTests - carriedFwdTests;
                    int totalPassed = testStatusDtSt.Tables[i].Rows[testStatusDtSt.Tables[i].Rows.Count - 1].Field<Int32>("Passed");
					int totalCondPassed = testStatusDtSt.Tables[i].Rows[testStatusDtSt.Tables[i].Rows.Count - 1].Field<Int32>("Conditionally Passed");
					int totalFailed = testStatusDtSt.Tables[i].Rows[testStatusDtSt.Tables[i].Rows.Count - 1].Field<Int32>("Failed");

					float percentCompleted = (float)(((float)totalPassed + (float)totalCondPassed + (float)totalFailed) / (float)total) * 100;
                        float percentPassed = (float)(((float)totalPassed + (float)totalCondPassed) / (float)total) * 100;

                        DataRow newRow = summaryTable.NewRow();
                        newRow[0] = releases[i];
                        newRow[1] = projects[i];
                        newRow[2] = totalTestCases;
                        newRow[3] = Math.Truncate(percentCompleted);
                        newRow[4] = Math.Truncate(percentPassed);
                        newRow[5] = releaseStatus[i];
                    string result = "";
                    foreach (string st in comments[i].Split(';'))
                    {
                        result =  result + st + "\n";
                    }
                    newRow[6] = comments[i];
                    summaryTable.Rows.Add(newRow);

                }

                CreateExcelReport excelReport = new CreateExcelReport();
                string excelFileName = "Test_Report" + DateTime.Now.ToString("ddmmyyyyhhmmss") + ".xlsx";
                //excelReport.generateTestExcelReport(testStatusSummaryDtset, testStatusValues, reportLocation, excelFileName);
                
                //excelFileName = "Test_Report_By_Severity" + DateTime.Now.ToString("ddmmyyyyhhmmss") + ".xlsx";

                //excelReport.generateTestBySeverityReport(testPriorityDtSt, reportLocation, excelFileName);
                excelReport.generateTestReport(summaryTable, testStatusDtSt, testPriorityDtSt, testStatusValues, reportLocation, excelFileName);
                //excelReport.colorCodeReleaseStatus(excelFileLocationToSave, excelFileName, releaseColorCode);

                if (arqResp)
                {
                    Bug bg = new Bug(tdc);

                    DataTable bugMasterTable = bg.bugsLinkedToTest(qcFldrNames);

                    excelFileName = "ARQ_Summary_" + DateTime.Now.ToString("ddmmyyyyhhmmss") + ".xlsx";
                    excelReport.generateArqExcelReport(bg.createBugReportBySubsystem(bugMasterTable), bg.createBugReportBySeverity(bugMasterTable),
                        reportLocation, excelFileName);
                }



            }//if (tdc.Connected) - ends here


            tdc.Disconnect();
            tdc.DisconnectProject();

    }



    }


}




