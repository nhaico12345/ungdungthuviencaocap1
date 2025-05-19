namespace ungdungthuviencaocap
{
	partial class bieudothongke
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea1 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
			System.Windows.Forms.DataVisualization.Charting.Legend legend1 = new System.Windows.Forms.DataVisualization.Charting.Legend();
			System.Windows.Forms.DataVisualization.Charting.Series series1 = new System.Windows.Forms.DataVisualization.Charting.Series();
			System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea2 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
			System.Windows.Forms.DataVisualization.Charting.Legend legend2 = new System.Windows.Forms.DataVisualization.Charting.Legend();
			System.Windows.Forms.DataVisualization.Charting.Series series2 = new System.Windows.Forms.DataVisualization.Charting.Series();
			this.chartSach = new System.Windows.Forms.DataVisualization.Charting.Chart();
			this.chartMuontra = new System.Windows.Forms.DataVisualization.Charting.Chart();
			((System.ComponentModel.ISupportInitialize)(this.chartSach)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.chartMuontra)).BeginInit();
			this.SuspendLayout();
			// 
			// chartSach
			// 
			chartArea1.Name = "ChartArea1";
			this.chartSach.ChartAreas.Add(chartArea1);
			this.chartSach.Dock = System.Windows.Forms.DockStyle.Left;
			legend1.Name = "Legend1";
			this.chartSach.Legends.Add(legend1);
			this.chartSach.Location = new System.Drawing.Point(0, 0);
			this.chartSach.Name = "chartSach";
			series1.ChartArea = "ChartArea1";
			series1.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
			series1.Legend = "Legend1";
			series1.Name = "Series1";
			this.chartSach.Series.Add(series1);
			this.chartSach.Size = new System.Drawing.Size(812, 774);
			this.chartSach.TabIndex = 0;
			this.chartSach.Text = "chart1";
			// 
			// chartMuontra
			// 
			chartArea2.Name = "ChartArea1";
			this.chartMuontra.ChartAreas.Add(chartArea2);
			this.chartMuontra.Dock = System.Windows.Forms.DockStyle.Right;
			legend2.Name = "Legend1";
			this.chartMuontra.Legends.Add(legend2);
			this.chartMuontra.Location = new System.Drawing.Point(818, 0);
			this.chartMuontra.Name = "chartMuontra";
			series2.ChartArea = "ChartArea1";
			series2.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Pie;
			series2.Legend = "Legend1";
			series2.Name = "Series1";
			this.chartMuontra.Series.Add(series2);
			this.chartMuontra.Size = new System.Drawing.Size(738, 774);
			this.chartMuontra.TabIndex = 1;
			this.chartMuontra.Text = "chartMuontra";
			// 
			// bieudothongke
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1556, 774);
			this.Controls.Add(this.chartMuontra);
			this.Controls.Add(this.chartSach);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			this.Name = "bieudothongke";
			this.Text = "Biểu đồ thống kê";
			((System.ComponentModel.ISupportInitialize)(this.chartSach)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.chartMuontra)).EndInit();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.DataVisualization.Charting.Chart chartSach;
		private System.Windows.Forms.DataVisualization.Charting.Chart chartMuontra;
	}
}