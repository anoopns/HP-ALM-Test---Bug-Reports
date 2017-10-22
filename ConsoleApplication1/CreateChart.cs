using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;

namespace ConsoleApplication1
{
    class CreateChart
    {
        public Bitmap Draw(Color bgColor, int width, int height,
       decimal[] vals)
        {
            // Create a new image and erase the background
            Bitmap bitmap = new Bitmap(width, height,
                                       PixelFormat.Format32bppArgb);
            Graphics graphics = Graphics.FromImage(bitmap);
            SolidBrush brush = new SolidBrush(bgColor);
            graphics.FillRectangle(brush, 0, 0, width, height);
            brush.Dispose();

            // Create brushes for coloring the pie chart
            SolidBrush[] brushes = new SolidBrush[10];
            brushes[0] = new SolidBrush(Color.Yellow);
            brushes[1] = new SolidBrush(Color.Green);
            brushes[2] = new SolidBrush(Color.Blue);
            brushes[3] = new SolidBrush(Color.Cyan);
            brushes[4] = new SolidBrush(Color.Magenta);
            brushes[5] = new SolidBrush(Color.Red);
            brushes[6] = new SolidBrush(Color.Black);
            brushes[7] = new SolidBrush(Color.Gray);
            brushes[8] = new SolidBrush(Color.Maroon);
            brushes[9] = new SolidBrush(Color.LightBlue);

            // Sum the inputs to get the total
            decimal total = 0.0m;
            foreach (decimal val in vals)
                total += val;

            // Draw the pie chart
            float start = 0.0f;
            float end = 0.0f;
            decimal current = 0.0m;
            for (int i = 0; i < vals.Length; i++)
            {
                current += vals[i];
                start = end;
                end = (float)(current / total) * 360.0f;
                graphics.FillPie(brushes[i % 10], 0.0f, 0.0f, width,
                                 height, start, end - start);
            }

            // Clean up the brush resources
            foreach (SolidBrush cleanBrush in brushes)
                cleanBrush.Dispose();

            return bitmap;
        }
    }
}
