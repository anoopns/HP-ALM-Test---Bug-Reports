using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Drawing;

namespace ConsoleApplication1
{
    class CreateHtmlPage
    {
        //Stearn writer to write to html format
        public static StreamWriter writer = null;
        //method to generate HTML wbe page by passing the value like test status count,test case name,pie chart location,header name of the tabular coloumn(status names),total test cases executed
        //location of html webpage
        //public static void saveHtmLPage(int[] values, string tCName, string imageFileLocation, string[] arrHeaderNames, int totalTestCaseExecutted, string strHtmlPageLocation)
        public static void saveHtmLPage(DataSet subsystemDataset, string[] testStatusValues, string strHtmlPageLocation, string imageFileLocation, List<string> graphNames)
        {
            //html web page location
            string filename = strHtmlPageLocation;
            //file stream 
            FileStream outputfile = null;
            outputfile = new FileStream(strHtmlPageLocation + "TestReport_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".html", 
                FileMode.OpenOrCreate, FileAccess.Write);
            writer = new StreamWriter(outputfile);

            try
            {
                writer.BaseStream.Seek(0, SeekOrigin.End);
                //write to html
                DoWrite("<HTML>");
                DoWrite("<HEAD>");
                DoWrite("<TITLE>");
                DoWrite("Test Summary Report");
                DoWrite("</TITLE>");
                DoWrite("<style>");
                DoWrite("table, td { white-space:pre }");
                DoWrite("th {background-color: #4CAF50; color: white;}");
                DoWrite("</style>");
                DoWrite("</HEAD>");
                DoWrite("<CENTER>");
                for(int i = 0; i< subsystemDataset.Tables.Count; i++)
                {
                    List<int> statusCount = new List<int>();

                    /* Copy the cound of each test status in a susystem table to generate the chart. 
                     * It doesn't copy from the first table as it is the summary table
                     *
                    */
            if (i != 0)
            {
                for (int col = 1; col < subsystemDataset.Tables[i].Columns.Count; col++)
                {
                    statusCount.Add(subsystemDataset.Tables[i].Rows[subsystemDataset.Tables[i].Rows.Count - 1].Field<Int32>(col));
                }

                Bitmap btMapImages = GeneratePieChart.pieChart(statusCount, testStatusValues, subsystemDataset.Tables[i].TableName);
                //Save that pie chart a particualar location
                string releaseName = subsystemDataset.Tables[i].TableName;
                string imageName = strHtmlPageLocation + releaseName.Replace(' ', '_') + ".png";
                btMapImages.Save(imageName, System.Drawing.Imaging.ImageFormat.Png);
                DoWrite("<img src=\"" + imageName + "\" alt='Big Boat'/> ");
                DoWrite("<BR></BR>");
            }

            DoWrite("<BR></BR>");
            DoWrite("<BR><H4>" + graphNames[i] + "</H4>");
            DoWrite(ConvertDataTableToHTML(subsystemDataset.Tables[i]));
            DoWrite("<BR></BR>");




        }

        DoWrite("</CENTER>");
        DoWrite("</BODY>");
        DoWrite("</HTML>");
        //close the writer
        writer.Close();
    }
    //catch if there is any exception
    catch (Exception ex)
    {
        Console.WriteLine("Exception GenerateCode = " + ex);
        // statusBar1.Text = "Error";
        outputfile = null;
        writer = null;
        //  return false;
    }

        }

        //writer to write in html format
        public static void DoWrite(String line)
        {
            writer.WriteLine(line);
            writer.Flush();
        }

        //method to form tabular format
        public static string ConvertDataTableToHTML(DataTable dt)
        {
            string html = "<table border='2' bordercolor='Black'>" +
                           "<tr bgcolor='Turquoise'>";
            //string html = "<table>";
            //add header row
            html += "<tr>";
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                html += "<td>\n" + dt.Columns[i].ColumnName + "</td>\n";
            }
                
            html += "</tr>\n";
            //add rows
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                html += "<tr>";
                for (int j = 0; j < dt.Columns.Count; j++)
                    html += "<td>\n" + dt.Rows[i][j].ToString() + "</td>\n";
                html += "</tr>\n";
            }
            html += "</table>";
            return html;
        }
    }
}
