using System;
using System.Collections.Generic;
using TDAPIOLELib;
using System.Data;

namespace ConsoleApplication1
{

	class TestData
	{
		private TDConnection tdc;
        /*A class that defines functions to create data tables for test reports.
         * */

		public TestData(TDConnection tdc)
		{
			this.tdc = tdc;
		}

        public DataTable createMasterTestTable(List<List<string>> releaseData)
        {
            /* Argument releaseData has release name, project name, cycle id's
             * Return a datatable with release, cycle id, status and priority 
             * this table is used to generate test reports by filtering with status and priority
             * */

            TestSetFactory testSetFactory = (TestSetFactory)tdc.TestSetFactory;
            TDFilter filter = (TDFilter)testSetFactory.Filter;
            DataTable masterTestTable = new DataTable();
            masterTestTable.Columns.Add("Release");
            masterTestTable.Columns.Add("Cycle");
            masterTestTable.Columns.Add("Status");
            masterTestTable.Columns.Add("Priority");


            foreach (List<string> release in releaseData)
            {
                if (release[0] != null & release[1] != null & release[2] != null)
                {
                    string releaseName = release[1];
                    string[] cycle_ids = release[2].Split(';');
                    foreach (string cycle_id in cycle_ids)
                    {
                        if (cycle_id.Trim(' ').Length > 0)
                        {
                            filter["CY_ASSIGN_RCYC"] = cycle_id;
                            //get all the list of test sets
                            try
                            {
                                List testSets = (List)testSetFactory.NewList(filter.Text);

                                if (testSets != null)
                                {
                                    foreach (TDAPIOLELib.TestSet testSet in testSets)
                                    {
                                        TSTestFactory testF = (TSTestFactory)testSet.TSTestFactory;
                                        List tests = (List)testF.NewList(" ");
                                        foreach (TDAPIOLELib.TSTest test in tests)
                                        {
                                            masterTestTable.Rows.Add(releaseName, cycle_id, test.Status, test["TS_USER_20"]);
                                        }
                                    }
                                }
                            }
                            catch(System.Runtime.InteropServices.COMException e)
                            {
                                Console.WriteLine("Enter valid test cycle id's");
                            }

                        }

                    }
                }
            }

            return masterTestTable;
        }


        public DataSet createTestReportByStatus(List<List<string>> releaseData, DataTable masterTestTable, string[] testStatusValues)
		{
            /* filter master table with status values and generate table for each release and add to a dataset which is returned 
             * by this function.
             * */
			DataSet testDtSt = new DataSet();
			
			foreach(List<string> list in releaseData)
			{
                /* list[0] = release name, list[1] = project name, list[2] = cycle id's
                 * */
				if (list[0] != null & list[1] != null & list[2] != null)
				{
					string projectName = list[1];
					string[] cycle_ids = list[2].Split(';');

					//Define columns 
					testDtSt.Tables.Add(projectName);
					testDtSt.Tables[projectName].Columns.Add("SubSystem");
					testDtSt.Tables[projectName].Columns.Add("Total Tests", typeof(Int32));
					foreach (string status in testStatusValues)
					{
						testDtSt.Tables[projectName].Columns.Add(status, typeof(Int32));
					}

					DataView dv = new DataView(masterTestTable);
					
					foreach (string cycle_id in cycle_ids)
					{
                        if(cycle_id.Trim(' ').Length > 0)
                        {
                            DataRow currentRow = testDtSt.Tables[projectName].NewRow();
                            dv.RowFilter = "Release = '" + projectName + "'" + "AND Cycle = '" + cycle_id + "'";
                            currentRow["SubSystem"] = Connect.getCycleName(tdc, cycle_id);
                            currentRow["Total Tests"] = dv.Count;
                            foreach (string status in testStatusValues)
                            {
                                dv.RowFilter = "Release = '" + projectName + "'" + "AND Cycle = '" + cycle_id + "'" +
                                    "AND Status = '" + status + "'";
                                currentRow[status] = dv.Count;
                            }
                            testDtSt.Tables[projectName].Rows.Add(currentRow);
                        }

					}
					DataRow lastRow = testDtSt.Tables[projectName].NewRow();
					lastRow["SubSystem"] = "Total";
					dv.RowFilter = "Release = '" + projectName + "'";
					lastRow["Total Tests"] = dv.Count;
					foreach (string status in testStatusValues)
					{
						dv.RowFilter = "Release = '" + projectName + "'" + "AND Status = '" + status + "'";
						lastRow[status] = dv.Count;
					}

					testDtSt.Tables[projectName].Rows.Add(lastRow);


				}
			}

			return testDtSt;
		}

		public DataSet createTestReportBySeverity(List<List<string>> releaseData, DataTable masterTestTable, string[] testStatusValues, string[] testPriorityValues)
		{
            /* filter master table with priority values and generate table for each release and add to a dataset which is returned 
             * by this function.
             * */
            DataSet testDtSt = new DataSet();

			foreach (List<string> release in releaseData)
			{
                /* list[0] = release name, list[1] = project name, list[2] = cycle id's
                * */
                if (release[0] != null & release[1] != null & release[2] != null)
				{
					string releaseName = release[1];
					string[] cycle_ids = release[2].Split(';');

					//Create a table to hold details of each release
					testDtSt.Tables.Add(releaseName);
					testDtSt.Tables[releaseName].Columns.Add("Priority");
					testDtSt.Tables[releaseName].Columns.Add("Total Tests");
					foreach (string status in testStatusValues)
					{
						testDtSt.Tables[releaseName].Columns.Add(status);
					}

					DataView dv = new DataView(masterTestTable);

					foreach (string priority in testPriorityValues)
					{
                        
						DataRow currentRow = testDtSt.Tables[releaseName].NewRow();
						dv.RowFilter = "Release = '" + releaseName + "'" + "AND Priority = '" + priority + "'";
						if (dv.Count != 0)
						{
							currentRow["Priority"] = priority;
							currentRow["Total Tests"] = dv.Count;
							foreach (string status in testStatusValues)
							{
								dv.RowFilter = "Release = '" + releaseName + "'" + "AND Priority = '" + priority + "'" +
									"AND Status = '" + status + "'";
								currentRow[status] = dv.Count;
							}
							testDtSt.Tables[releaseName].Rows.Add(currentRow);

						}
						
						

					}
					DataRow lastRow = testDtSt.Tables[releaseName].NewRow();
					lastRow["Priority"] = "Total";
					dv.RowFilter = "Release = '" + releaseName + "'";
					lastRow["Total Tests"] = dv.Count;
					foreach (string status in testStatusValues)
					{
						dv.RowFilter = "Release = '" + releaseName + "'" + "AND Status = '" + status + "'";
						lastRow[status] = dv.Count;
					}

					testDtSt.Tables[releaseName].Rows.Add(lastRow);


				}
			}

			return testDtSt;
		}




	}

}

