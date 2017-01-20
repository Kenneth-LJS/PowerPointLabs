using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Drawing;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FillFormat : ObjectFormat
    {
        static private int index = 0;
        readonly Microsoft.Office.Interop.PowerPoint.FillFormat style;

        public FillFormat(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            this.displayText = "Fill " + (index++).ToString();
         
            this.displayImage = Utils.Graphics.CreateImageFromShape(shape);
            //this.style = new Microsoft.Office.Interop.PowerPoint.FillFormat(style);
            this.style = null;
            //System.Drawing.Bitmap b = new System.Drawing.Bitmap(DISPLAY_IMAGE_SIZE.Width, DISPLAY_IMAGE_SIZE.Height);
            //Graphics g = Graphics.FromImage(b);
            //g.FillRectangle(new SolidBrush(System.Drawing.ColorTranslator.FromOle(style.ForeColor.RGB)), 0, 0, b.Width, b.Height);
            //this.displayImage = b;
        }

        public override void ApplyTo(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            shape.Fill.ForeColor = style.ForeColor;
            shape.Fill.BackColor = style.BackColor;
        }
    }
}
