using System;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data;
using System.Drawing;

namespace ConsoleApplication1
{
    class CreateExcelReport
    {
        Microsoft.Office.Interop.Excel.Application excel;
        Microsoft.Office.Interop.Excel.Workbook excelworkBook;
        Microsoft.Office.Interop.Excel.Worksheet excelSheet;
        Microsoft.Office.Interop.Excel.Range excelCellrange;

        public void generateArqExcelReport(DataSet arqCIDtSt, DataSet arqSevDtSt , string excelLocationToSave, string filename)
        {
            string excelFileToSave = excelLocationToSave + filename;
            try
            {
                // Start Excel and get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();

                // for making Excel visible
                excel.Visible = false;
                excel.DisplayAlerts = false;

                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Add(Type.Missing);

                // Work sheet
                for (int count = 0; count < arqCIDtSt.Tables.Count; count++)
                {
                    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add();
                    excelSheet.Name = arqCIDtSt.Tables[count].TableName;
                    //excelSheet.Name = "Project_" + count;
                    excelSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                    excelSheet.PageSetup.Zoom = false;
                    excelSheet.PageSetup.FitToPagesWide = 1;
                    excelSheet.PageSetup.FitToPagesTall = 1;

                    int rowStart = 2;
                    int colStart = 2;

                    fillExcelSheetWithDtTblValues(excelSheet, arqCIDtSt.Tables[count], rowStart, colStart);

                    rowStart = arqCIDtSt.Tables[count].Rows.Count + 4;
                    colStart = 2;

                    fillExcelSheetWithDtTblValues(excelSheet, arqSevDtSt.Tables[count], rowStart, colStart);

                }
                //now save the workbook and exit Excel

                excelworkBook.SaveAs(excelFileToSave);
                excelworkBook.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

            }
            finally
            {
                excelSheet = null;
                excelCellrange = null;
                excelworkBook = null;
            }

        }//End of function generateExcelReport

        public Excel.Worksheet fillExcelSheetWithDtTblValues(Microsoft.Office.Interop.Excel.Worksheet excelSheet, DataTable table, 
            int rowStart, int colStart)
        {
            int rowIndex = rowStart;
            int colIndex = colStart;

            foreach (DataColumn col in table.Columns)
            {
                excelSheet.Cells[rowIndex, colIndex] = col.ColumnName;
                colIndex++;
            }


            //To format the first row of the table
            excelCellrange = excelSheet.Range[excelSheet.Cells[rowStart, colStart],
                excelSheet.Cells[rowStart, colIndex-1]];
            FormattingExcelCells(excelCellrange, "#000099", System.Drawing.Color.White, true);
            excelCellrange.EntireColumn.AutoFit();
            excelCellrange.RowHeight = 30;
            Excel.Borders border = excelCellrange.Borders;
            border.LineStyle = Excel.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            colIndex = colStart;
            rowIndex = rowStart + 1;
            foreach (DataRow row in table.Rows)
            {
                foreach(DataColumn col in table.Columns)
                {
                    excelSheet.Cells[rowIndex, colIndex] = row[col].ToString();
                    colIndex++;
                }
                // now we resize the row
                excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex, colStart],
                    excelSheet.Cells[rowIndex, colIndex-1]];
                excelCellrange.EntireColumn.AutoFit();
                excelCellrange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                border = excelCellrange.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;

                colIndex = colStart;
                rowIndex++;
            }
            //To format the last row of the table except the summary table
            if (!table.TableName.Contains("Test_Summary"))
            {
                excelCellrange = excelSheet.Range[excelSheet.Cells[rowIndex-1, colStart],
                    excelSheet.Cells[rowIndex-1, table.Columns.Count+1]];
                FormattingExcelCells(excelCellrange, "#000099", System.Drawing.Color.White, true);
                border = excelCellrange.Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                border.Weight = 2d;
            }

            return excelSheet;

        }

        public void generateTestReport(DataTable summaryTbl, DataSet testCIDtSt, DataSet testSevDtSt, string[] testStatusValues, 
            string excelLocationToSave, string excelFileName)
        {

            try
            {
                // Start Excel and get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();

                // for making Excel visible
                excel.Visible = false;
                excel.DisplayAlerts = false;

                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Add(Type.Missing);

                int rowstart;
                int colstart;



                // Work sheet
                for (int count = testCIDtSt.Tables.Count - 1; count >= 0; count--)
                {
                    excelSheet = (Excel.Worksheet)excelworkBook.Worksheets.Add();
                    excelSheet.Name = testCIDtSt.Tables[count].TableName;
                    excelSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
                    excelSheet.PageSetup.Zoom = false;
                    excelSheet.PageSetup.FitToPagesWide = 1;
                    excelSheet.PageSetup.FitToPagesTall = 1;
                    rowstart = 17;
                    colstart = 2;

                    fillExcelSheetWithDtTblValues(excelSheet, testCIDtSt.Tables[count], rowstart, colstart);

                    //To start the next table leaving 2 rows after the first table
                    rowstart = testCIDtSt.Tables[count].Rows.Count + rowstart + 3;

                    fillExcelSheetWithDtTblValues(excelSheet, testSevDtSt.Tables[count], rowstart, 2);

                    if (!testCIDtSt.Tables[count].TableName.Contains("Test_Summary"))
                    {
                        /*Excel.ChartObjects chartObjs = (Excel.ChartObjects)excelSheet.ChartObjects(Type.Missing);
                        Excel.ChartObject chartObj = chartObjs.Add(100, 20, 300, 200);
                        Excel.Chart xlChart = chartObj.Chart;
                        Excel.Range chartRage = excelSheet.Range["D17:K17,D28:K28"];
                        xlChart.ChartType = Excel.XlChartType.xlPie;
                        xlChart.SetSourceData(chartRage, Type.Missing);*/

                        List<int> totals = new List<int>();
                        for (int col = 1; col < testCIDtSt.Tables[count].Columns.Count; col++)
                        {
                            totals.Add(
                                testCIDtSt.Tables[count].Rows[testCIDtSt.Tables[count].Rows.Count - 1].Field<Int32>(col));
                        }
                        Bitmap btMapImages = GeneratePieChart.pieChart(totals, testStatusValues, testCIDtSt.Tables[count].TableName);
                        //Save that pie chart a particualar location
                        string releaseName = testCIDtSt.Tables[count].TableName;
                        string imageName = excelLocationToSave + releaseName.Replace(' ', '_') + ".png";
                        btMapImages.Save(imageName, System.Drawing.Imaging.ImageFormat.Png);
                        excelSheet.Shapes.AddPicture(imageName, Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 100, 10, 350, 200);

                    }



                }

                excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add();
                excelSheet.Name = summaryTbl.TableName;
                excelSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                excelSheet.PageSetup.Zoom = false;
                excelSheet.PageSetup.FitToPagesWide = 1;
                excelSheet.PageSetup.FitToPagesTall = 1;

                rowstart = 2;
                colstart = 2;
                fillExcelSheetWithDtTblValues(excelSheet, summaryTbl, rowstart, colstart);

                //now save the workbook and exit Excel
                string fileLocation = excelLocationToSave + excelFileName;
                excelworkBook.SaveAs(fileLocation);
                excelworkBook.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);

            }
            finally
            {
                excelSheet = null;
                excelCellrange = null;
                excelworkBook = null;
            }

        }//End of function "generateTestExcelReport"


        public void generateTestExcelReport(DataSet dtSet, string[] testStatusValues, string excelLocationToSave, string excelFileName)
        {

			try
            {
                // Start Excel and get Application object.
                excel = new Microsoft.Office.Interop.Excel.Application();

                // for making Excel visible
                excel.Visible = false;
                excel.DisplayAlerts = false;

                // Creation a new Workbook
                excelworkBook = excel.Workbooks.Add(Type.Missing);


                // Work sheet
                for (int count = dtSet.Tables.Count -1; count >=0 ; count--)
                {
                    excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add();
                    excelSheet.Name = dtSet.Tables[count].TableName;
                    excelSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
                    excelSheet.PageSetup.Zoom = false;
                    excelSheet.PageSetup.FitToPagesWide = 1;
                    excelSheet.PageSetup.FitToPagesTall = 1;
                    
                    int rowcount, rowstart;
                    /*
                     * Determine the start row number in the sheet.
                     * Summary table (count =0) starts at row 0 and rest at 17 to leave space for the pie chart
                     */
                    if (count == 0)
                    {
                        rowcount = 2;
                        rowstart = 2;
                    }
                    else
                    {
                        rowcount = 17;
                        rowstart = 17;
                    }


                    for (int i = 0; i < dtSet.Tables[count].Columns.Count; i++)
                    {
                        excelSheet.Cells[rowstart, i + 2] = dtSet.Tables[count].Columns[i].ColumnName;
                        excelSheet.Cells.Font.Color = System.Drawing.Color.Black;

                    }

                    //To format the first row of the table
                    excelCellrange = excelSheet.Range[excelSheet.Cells[rowstart, 2],
                        excelSheet.Cells[rowstart, dtSet.Tables[count].Columns.Count + 1]];
                    FormattingExcelCells(excelCellrange, "#000099", System.Drawing.Color.White, true);
                    excelCellrange.EntireColumn.AutoFit();
                    excelCellrange.RowHeight = 30;

                    foreach (DataRow datarow in dtSet.Tables[count].Rows)
                    {
                        rowcount += 1;
                        for (int j = 0; j < dtSet.Tables[count].Columns.Count; j++)
                        {

                            excelSheet.Cells[rowcount, j + 2] = datarow[j].ToString();
                        }

                        // now we resize the columns
                        excelCellrange = excelSheet.Range[excelSheet.Cells[rowstart, 2], 
                            excelSheet.Cells[rowcount, dtSet.Tables[count].Columns.Count + 1]];
                        excelCellrange.EntireColumn.AutoFit();
                        Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
                        border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
                        border.Weight = 2d;

                    }
                    //To format the last row of the table except the summary table
                    if (!dtSet.Tables[count].TableName.Contains("Test_Summary"))
                    {
                        excelCellrange = excelSheet.Range[excelSheet.Cells[rowcount, 2],
                            excelSheet.Cells[rowcount, dtSet.Tables[count].Columns.Count + 1]];
                        FormattingExcelCells(excelCellrange, "#000099", System.Drawing.Color.White, true);
                    }


                    if (count != 0)
                    {
                        /*Excel.ChartObjects chartObjs = (Excel.ChartObjects)excelSheet.ChartObjects(Type.Missing);
                        Excel.ChartObject chartObj = chartObjs.Add(100, 20, 300, 200);
                        Excel.Chart xlChart = chartObj.Chart;
                        Excel.Range chartRage = excelSheet.Range["D17:K17,D28:K28"];
                        xlChart.ChartType = Excel.XlChartType.xlPie;
                        xlChart.SetSourceData(chartRage, Type.Missing);*/

                        List<int> totals = new List<int>();
                        for (int col = 1; col < dtSet.Tables[count].Columns.Count; col++)
                        {
                            totals.Add(
                                dtSet.Tables[count].Rows[dtSet.Tables[count].Rows.Count - 1].Field<Int32>(col));
                        }
                        Bitmap btMapImages = GeneratePieChart.pieChart(totals, testStatusValues, dtSet.Tables[count].TableName);
                        //Save that pie chart a particualar location
                        string releaseName = dtSet.Tables[count].TableName;
                        string imageName = excelLocationToSave + releaseName.Replace(' ', '_') + ".png";
                        btMapImages.Save(imageName, System.Drawing.Imaging.ImageFormat.Png);
                        excelSheet.Shapes.AddPicture(imageName, Microsoft.Office.Core.MsoTriState.msoFalse,
                            Microsoft.Office.Core.MsoTriState.msoTrue, 100, 10, 350, 200);

                    }

                }
                //now save the workbook and exit Excel

                string fileLocation = excelLocationToSave + excelFileName;
                excelworkBook.SaveAs(fileLocation);
                excelworkBook.Close();
                excel.Quit();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                
            }
            finally
            {
                excelSheet = null;
                excelCellrange = null;
                excelworkBook = null;
            }

        }//End of function "generateTestExcelReport"

		public void generateTestBySeverityReport(DataSet dtSet, string excelLocationToSave, string excelFileName)
		{

			try
			{
				// Start Excel and get Application object.
				excel = new Microsoft.Office.Interop.Excel.Application();

				// for making Excel visible
				excel.Visible = false;
				excel.DisplayAlerts = false;

				// Creation a new Workbook
				excelworkBook = excel.Workbooks.Add(Type.Missing);

				string[] testStatusNames = new string[dtSet.Tables[0].Columns.Count - 2];
				for (int i = 0; i < dtSet.Tables[0].Columns.Count - 1; i++)
				{
					if (i != 0)
					{
						testStatusNames[i - 1] = dtSet.Tables[0].Columns[i].ColumnName;
					}

				}


                // Work sheet
                for (int count = dtSet.Tables.Count - 1; count >= 0; count--)
                {
					excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelworkBook.Worksheets.Add();
					excelSheet.Name = dtSet.Tables[count].TableName;
					excelSheet.PageSetup.Orientation = Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;
					excelSheet.PageSetup.Zoom = false;
					excelSheet.PageSetup.FitToPagesWide = 1;
					excelSheet.PageSetup.FitToPagesTall = 1;

					int rowcount, rowstart;
					/*
                     * Determine the start row number in the sheet.
                     */
						rowcount = 2;
						rowstart = 2;


					for (int i = 0; i < dtSet.Tables[count].Columns.Count; i++)
					{
						excelSheet.Cells[rowstart, i + 2] = dtSet.Tables[count].Columns[i].ColumnName;
						excelSheet.Cells.Font.Color = System.Drawing.Color.Black;

					}

					//To format the first row of the table
					excelCellrange = excelSheet.Range[excelSheet.Cells[rowstart, 2],
						excelSheet.Cells[rowstart, dtSet.Tables[count].Columns.Count + 1]];
					FormattingExcelCells(excelCellrange, "#000099", System.Drawing.Color.White, true);
					excelCellrange.EntireColumn.AutoFit();
					excelCellrange.RowHeight = 30;

					foreach (DataRow datarow in dtSet.Tables[count].Rows)
					{
						rowcount += 1;
						for (int j = 0; j < dtSet.Tables[count].Columns.Count; j++)
						{

							excelSheet.Cells[rowcount, j + 2] = datarow[j].ToString();
						}

						// now we resize the columns
						excelCellrange = excelSheet.Range[excelSheet.Cells[rowstart, 2],
							excelSheet.Cells[rowcount, dtSet.Tables[count].Columns.Count + 1]];
						excelCellrange.EntireColumn.AutoFit();
						Microsoft.Office.Interop.Excel.Borders border = excelCellrange.Borders;
						border.LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
						border.Weight = 2d;

					}
					//To format the last row of the table except the summary table
						excelCellrange = excelSheet.Range[excelSheet.Cells[rowcount, 2],
							excelSheet.Cells[rowcount, dtSet.Tables[count].Columns.Count + 1]];
						FormattingExcelCells(excelCellrange, "#000099", System.Drawing.Color.White, true);

				}
				//now save the workbook and exit Excel

				string fileLocation = excelLocationToSave + excelFileName;
				excelworkBook.SaveAs(fileLocation);
				excelworkBook.Close();
				excel.Quit();
			}
			catch (Exception ex)
			{
				Console.WriteLine(ex.Message);

			}
			finally
			{
				excelSheet = null;
				excelCellrange = null;
				excelworkBook = null;
			}

		}//End of function "generateTestExcelReport"


		/// <summary>
		/// FUNCTION FOR FORMATTING EXCEL CELLS
		/// </summary>
		/// <param name="range"></param>
		/// <param name="HTMLcolorCode"></param>
		/// <param name="fontColor"></param>
		/// <param name="IsFontbool"></param>
		public void FormattingExcelCells(Microsoft.Office.Interop.Excel.Range range, string HTMLcolorCode, System.Drawing.Color fontColor, bool IsFontbool)
        {
            range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(HTMLcolorCode);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor);
            if (IsFontbool == true)
            {
                range.Font.Bold = IsFontbool;
            }
        }

        public void colorCodeReleaseStatus(string excelLocation, string excelFileName, Dictionary<string, string> releaseColorCode)
        {
            excel = new Microsoft.Office.Interop.Excel.Application();
            string fileName = excelLocation + excelFileName;
            excelworkBook = excel.Workbooks.Open(@fileName, 0, false, 5,"", "", true, 
                Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            excelSheet = (Excel.Worksheet)excelworkBook.Worksheets.get_Item(1);
            excelCellrange = excelSheet.UsedRange;

            int rw = excelCellrange.Rows.Count + 1;

            for (int rCnt = 3; rCnt <= rw; rCnt++)
            {

                if (releaseColorCode.ContainsKey((string)(excelSheet.Cells[rCnt,3]).Value))
                {
                    excelCellrange = excelSheet.Range[excelSheet.Cells[rCnt, 7],
                        excelSheet.Cells[rCnt, 7]];
                    FormattingExcelCells(excelCellrange, releaseColorCode[(string)(excelSheet.Cells[rCnt, 3]).Value], 
                        System.Drawing.Color.White, true);
                }

            }

            excel.DisplayAlerts = false;
            excelworkBook.SaveAs(fileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, 
                true, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, 
                Microsoft.Office.Interop.Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            excelworkBook.Close(true, null, null);
            excel.Quit();

            Marshal.ReleaseComObject(excelSheet);
            Marshal.ReleaseComObject(excelworkBook);
            Marshal.ReleaseComObject(excel);
        }
    }
}

