using System.Data;
using System;
using TDAPIOLELib;
using System.Collections.Generic;

namespace ConsoleApplication1
{
    class Bug
    {
        public TDConnection tdc;

        //To track project values
        HashSet<string> projects = new HashSet<string>();
        //To track qc folder names (referred here as releases)
        HashSet<string> releases = new HashSet<string>();

        public Bug(TDConnection tdc)
        {
            this.tdc = tdc;
        }

        public string processQcFldrString(string qcFldrName)
        {
            /*
             * Returns the child folder name
             */
            string[] processed = qcFldrName.Split('\\');
            return processed[processed.Length-1];
        }

        public DataTable bugsLinkedToTest(string[] testFldrNames)
        {
            /*
             * Uses LinkFactory to find bugs linked a test run
             */

            LinkFactory linkF;
            ILinkable linkableTest;
            TestSetFactory tsFactory = (TestSetFactory)tdc.TestSetFactory;
            TestSetTreeManager tsTreeMgr = (TestSetTreeManager)tdc.TestSetTreeManager;
            BugFactory bgf = (BugFactory)tdc.BugFactory;


            TDFilter bugFilter = (TDFilter)bgf.Filter;

            TDFilter testFilter = (TDFilter)tsFactory.Filter;

            HashSet<string> bugIds = new HashSet<string>();

            //Table to hold all bugs 
            DataTable bugMasterTable = new DataTable();
            bugMasterTable.Columns.Add("release");
            bugMasterTable.Columns.Add("project");
            bugMasterTable.Columns.Add("subsystem");
            bugMasterTable.Columns.Add("status");
            bugMasterTable.Columns.Add("severity");

            //To remove all existing values if any
            projects.Clear();
            releases.Clear();

            foreach (string qcFldr in testFldrNames)
            {
                if (qcFldr.Trim(' ').Length > 0)
                {
                    string release = processQcFldrString(qcFldr);
                    releases.Add(release);

                    TestSetFolder tsFolder = (TestSetFolder)tsTreeMgr.get_NodeByPath(qcFldr);
                    List tsList = tsFolder.FindTestSets("", false, null);

                    foreach (TDAPIOLELib.TestSet testSt in tsList)
                    {
                        TSTestFactory testF = (TSTestFactory)testSt.TSTestFactory;
                        List tests = (List)testF.NewList(" ");

                        foreach (TDAPIOLELib.TSTest test in tests)
                        {

                            linkableTest = (ILinkable)test;
                            linkF = linkableTest.BugLinkFactory;
                            if (linkableTest.HasLinkage)
                            {
                                testFilter = (TDFilter)linkF.Filter;
                                testFilter["LN_ENTITY_ID"] = test.ID;
                                List links = (List)linkF.NewList("");

                                foreach (TDAPIOLELib.Link ln in links)
                                {
                                    bugFilter["BG_BUG_ID"] = Convert.ToString(ln["LN_BUG_ID"]);
                                    List bugList = (List)bgf.NewList(bugFilter.Text);
                                    foreach (TDAPIOLELib.Bug bg in bugList)
                                    {
                                        //Add details of only unique bug id's
                                        if (!bugIds.Contains(Convert.ToString(bg.ID)))
                                        {
                                            projects.Add(bg["BG_PROJECT"]);
                                            bugMasterTable.Rows.Add(release, bg["BG_PROJECT"], bg["BG_USER_03"], bg.Status, bg["BG_SEVERITY"]);
                                            bugIds.Add(Convert.ToString(bg.ID));
                                        }
                                    }
                                }
                            }


                        }
                    }

                }
            }
            return bugMasterTable;
        }


        /*public DataTable createMasterBugDtSt(string[] releases)
        {
            BugFactory bgf = (BugFactory)tdc.BugFactory;
            

            TDFilter filter = (TDFilter)bgf.Filter;

            //Table to hold all bugs 
            DataTable bugMasterTable = new DataTable();
            bugMasterTable.Columns.Add("project");
            bugMasterTable.Columns.Add("subsystem");
            bugMasterTable.Columns.Add("status");
            bugMasterTable.Columns.Add("severity");

            //To remove all existing values if any
            projects.Clear();

            foreach (string release in releases)
            {
                if (release.Trim(' ').Length > 0)
                {
                    string filterValue = "\"" + release + "\"";
                    filter["BG_USER_12"] = filterValue;
                    List bugList = (List)bgf.NewList(filter.Text);

                    foreach (TDAPIOLELib.Bug bug in bugList)
                    {
                        
                        projects.Add(bug["BG_PROJECT"]);
                        bugMasterTable.Rows.Add(bug["BG_PROJECT"], bug["BG_USER_03"], bug.Status, bug["BG_SEVERITY"]);
                    }
                }
            }

            return bugMasterTable;
        }*/

        public DataSet createBugReportBySubsystem(DataTable bugMasterTable)
        {
            /*
             * Create a DataSet with tables for each qc folder (release) 
             * Each table has Bug status in sub system order
             */

            DataSet bugDtSt = new DataSet();

            DataView dv = new DataView(bugMasterTable);

            foreach(string release in releases)
            {
                DataTable bugDtTbl = new DataTable(release);
                dv.RowFilter = "release = '" + release + "'";

                DataTable releaseTable = new DataTable();
                releaseTable = dv.ToTable();

                HashSet<string> subSystemValues = new HashSet<string>();
                HashSet<string> statusValues = new HashSet<string>();

                foreach (DataRow row in releaseTable.Rows)
                {
                    subSystemValues.Add(row.Field<string>("subsystem"));
                }

                foreach (DataRow row in releaseTable.Rows)
                {
                    statusValues.Add(row.Field<string>("status"));
                }

                bugDtTbl.Columns.Add("SubSystem");

                foreach (string status in statusValues)
                {
                    bugDtTbl.Columns.Add(status);
                }
                bugDtTbl.Columns.Add("Total");

                DataView dv1 = new DataView(releaseTable);


                foreach (string subsystem in subSystemValues)
                {
                    DataRow currentRow = bugDtTbl.NewRow();
                    currentRow["SubSystem"] = subsystem;


                    foreach (string status in statusValues)
                    {
                        dv1.RowFilter = "subsystem = '" + subsystem + "'" + "AND status = '" + status + "'";
                        currentRow[status] = dv1.Count;
                    }
                    dv1.RowFilter = "subsystem = '" + subsystem + "'";
                    currentRow["Total"] = dv1.Count;
                    bugDtTbl.Rows.Add(currentRow);


                }
                DataRow lastRow = bugDtTbl.NewRow();
                lastRow["SubSystem"] = "Total";
                foreach (string status in statusValues)
                {
                    dv1.RowFilter = "status = '" + status + "'";
                    lastRow[status] = dv1.Count;
                }

                lastRow["Total"] = releaseTable.Rows.Count;
                bugDtTbl.Rows.Add(lastRow);
                releaseTable.Reset();

                bugDtSt.Tables.Add(bugDtTbl);
            }


            return bugDtSt;
        }

        public DataSet createBugReportBySeverity(DataTable bugMasterTable)
        {
            /*
             * Create a DataSet with tables for each qc folder (release) 
             * Each table has Bug status in Bug Severity order
             */
            DataSet bugDtSt = new DataSet();

            DataView dv = new DataView(bugMasterTable);

            foreach (string release in releases)
            {
                DataTable bugDtTbl = new DataTable(release);
                dv.RowFilter = "release = '" + release + "'";

                DataTable releaseTable = new DataTable();
                releaseTable = dv.ToTable();

                HashSet<string> severityValues = new HashSet<string>();
                HashSet<string> statusValues = new HashSet<string>();

                foreach (DataRow row in releaseTable.Rows)
                {
                    severityValues.Add(row.Field<string>("severity"));
                }

                foreach (DataRow row in releaseTable.Rows)
                {
                    statusValues.Add(row.Field<string>("status"));
                }

                bugDtTbl.Columns.Add("severity");

                foreach (string status in statusValues)
                {
                    bugDtTbl.Columns.Add(status);
                }
                bugDtTbl.Columns.Add("Total");

                DataView dv1 = new DataView(releaseTable);


                foreach (string severity in severityValues)
                {
                    DataRow currentRow = bugDtTbl.NewRow();
                    currentRow["severity"] = severity;


                    foreach (string status in statusValues)
                    {
                        dv1.RowFilter = "severity = '" + severity + "'" + "AND status = '" + status + "'";
                        currentRow[status] = dv1.Count;
                    }
                    dv1.RowFilter = "severity = '" + severity + "'";
                    currentRow["Total"] = dv1.Count;
                    bugDtTbl.Rows.Add(currentRow);


                }
                DataRow lastRow = bugDtTbl.NewRow();
                lastRow["severity"] = "Total";
                foreach (string status in statusValues)
                {
                    dv1.RowFilter = "status = '" + status + "'";
                    lastRow[status] = dv1.Count;
                }

                lastRow["Total"] = releaseTable.Rows.Count;
                bugDtTbl.Rows.Add(lastRow);
                releaseTable.Reset();

                bugDtSt.Tables.Add(bugDtTbl);
            }


            return bugDtSt;
        }

    }


}
