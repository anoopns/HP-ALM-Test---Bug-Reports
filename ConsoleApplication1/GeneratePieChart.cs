using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Drawing;
using System.Collections.Generic;
using System.Drawing.Drawing2D;

namespace ConsoleApplication1
{
    class GeneratePieChart
    {

        //This method will generate Pie Chart
        public static Bitmap pieChart(List<int> values, string[] valuesName, string chartTitle)
        {
            //Bitmap objBitmap = new Bitmap(400, 200);
            Bitmap objBitmap = new Bitmap(350, 200);
            try
            {
                //calculate total


                //Get the tOtal for the pie chart
                float total = values[0];

                Graphics objGraphics = Graphics.FromImage(objBitmap);
                objGraphics.Clear(Color.WhiteSmoke);

                objGraphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                objGraphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;

                objGraphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;

                objGraphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;


                objGraphics.DrawString(chartTitle, new Font("Tahoma", 10), Brushes.Black, new PointF(5, 5));

                //PointF symbolLeg = new PointF(210, 32);
                //PointF descLeg = new PointF(230, 30);
                PointF symbolLeg = new PointF(190, 26);
                PointF descLeg = new PointF(210, 24);
                //now draw the pie chart
                float sglCurrentAngle = 0;
                float sglTotalAngle = 0;

                for (int i = 1; i < values.Count; i++)
                {
                    if (valuesName != null)
                    {
                        float value = ((float)values[i]) / total;
                        sglCurrentAngle = value * 360;
                        //objGraphics.FillPie(new SolidBrush(GetColor(i)), 50, 40, 150, 150, sglTotalAngle, sglCurrentAngle);
                        objGraphics.FillPie(new SolidBrush(GetColor(i)), 30, 30, 140, 140, sglTotalAngle, sglCurrentAngle);
                        //objGraphics.DrawPie(Pens.Black, 50, 40, 150, 150, sglTotalAngle, sglCurrentAngle);
                        objGraphics.DrawPie(Pens.Black, 30, 30, 140, 140, sglTotalAngle, sglCurrentAngle);

                        sglTotalAngle += sglCurrentAngle;
                    }
                }

                //Draw the Rectangular value names and with color

                for (int i=1; i< values.Count; i++)
                {
                    //need to findout the percentage
                    if (valuesName != null)
                    {
                        float percenTage = (float)(((float)values[i]) / (float)total) * 100;
                        Int16 intPercenTage = Convert.ToInt16(percenTage);
                        objGraphics.FillRectangle(new SolidBrush(GetColor(i)), symbolLeg.X, symbolLeg.Y, 20, 10);
                        //objGraphics.DrawString(valuesName[i-1].ToString() + "[" + intPercenTage + "%]", new Font("Tahoma", 8), Brushes.Black, descLeg);
                        objGraphics.DrawString(valuesName[i - 1].ToString() + "[" + Math.Round((decimal)percenTage,2) + "%]", new Font("Tahoma", 8), Brushes.Black, descLeg);
                        symbolLeg.Y += 15;
                        descLeg.Y += 15;
                    }

                }


            }
            catch (Exception ex)
            {


            }
            return objBitmap;
        }

        //method to get color code
        public static Color GetColor(int colorCode)
        {

            Color objColor = Color.Red;
            try
            {
                switch (colorCode)
                {

                    case 1:
                        //code for passed
                        objColor = Color.Green;
                        break;

                    case 2:
                        //code for cond' passed
                        objColor = Color.GreenYellow;
                        break;

                    case 3:
                        //code for failed
                        objColor = Color.Red;
                        break;

                    case 4:
                        //code for No Run
                        objColor = Color.Blue;
                        break;

                    case 5:
                        //code for Descoped
                        objColor = Color.Gray;
                        break;

                    case 6:
                        //code for blocked
                        objColor = Color.Black;
                        break;

                    case 7:
                        //code for not completed
                        objColor = Color.Orange;
                        break;

                    case 8:
                        //code for Carried Forward
                        objColor = Color.Maroon;
                        break;
                    case 9:

                        objColor = Color.Indigo;
                        break;
                    case 10:

                        objColor = Color.Tan;
                        break;
                    case 11:

                        objColor = Color.Khaki;
                        break;
                    case 12:

                        objColor = Color.Salmon;
                        break;
                    case 13:

                        objColor = Color.Olive;
                        break;
                    default:
                        objColor = Color.Turquoise;
                        break;

                }
            }
            catch (Exception ex)
            {
                //
            }
            return objColor;
        }
    }
}
