using DevExpress.XtraDashboardLayout;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ungdungthuviencaocap
{
    public class ArtanPictureBox: PictureBox
    {
		// Fields
		private int boderSize = 2;
		private int boderRadius = 40;
		private float gradientAngle = 90F;
		private Color boderGradientTop = Color.DodgerBlue;
		private Color boderGradientBottom = Color.DodgerBlue;

		// Constructor
		public ArtanPictureBox()
		{
			SizeMode = PictureBoxSizeMode.StretchImage;
			Size = new Size(120, 120);
		}

		public int BoderSize
		{
			get {return boderSize;}
			set {boderSize = value;Invalidate();}
		}

		public int BoderRadius
		{
			get { return boderRadius; }
			set { boderRadius = value; Invalidate(); }
		}

		public float GradientAngle
		{
			get { return gradientAngle; }
			set { gradientAngle = value; Invalidate(); }
		}

		public Color BoderGradientTop
		{
			get { return boderGradientTop; }
			set { boderGradientTop = value; Invalidate(); }
		}

		public Color BoderGradientBottom
		{
			get { return boderGradientBottom; }
			set { boderGradientBottom = value; Invalidate(); }
		}

		// Overriden method
		protected override void OnResize(EventArgs e)
		{
			base.OnResize(e);
			Size = new Size(Width, Width);
		}

		protected override void OnPaint(PaintEventArgs pe)
		{
			base.OnPaint(pe);
			// Fields
			var graphics = pe.Graphics;
			var rectangleSmooth = Rectangle.Inflate(ClientRectangle, -1, -1);
			var rectangleBorder = Rectangle.Inflate(rectangleSmooth, -boderSize, -boderSize);
			var smoothSize = boderRadius > 0? boderSize * 3 : 1;

			using (var borderGradientColor = new LinearGradientBrush(rectangleBorder, boderGradientTop, boderGradientBottom, gradientAngle))
			using (var pathRegion = new GraphicsPath())
			using (var penSmooth = new Pen(Parent.BackColor, smoothSize))
			using (var penBoder = new Pen(borderGradientColor, smoothSize))
			{
				pathRegion.AddEllipse(rectangleSmooth);
				Region = new Region(pathRegion);
				graphics.SmoothingMode = SmoothingMode.AntiAlias;
				graphics.DrawEllipse(penSmooth, rectangleSmooth);

				if (boderSize > 0)
					graphics.DrawEllipse(penBoder, rectangleBorder);
			}
			Size = new Size(Width, Width);
		}
	}
}
