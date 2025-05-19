namespace ungdungthuviencaocap
{
	partial class bangthongke
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
			this.components = new System.ComponentModel.Container();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(bangthongke));
			this.ribbon = new DevExpress.XtraBars.Ribbon.RibbonControl();
			this.barButtonItem1 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem2 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem3 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem4 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem5 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem6 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem7 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem8 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem9 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem10 = new DevExpress.XtraBars.BarButtonItem();
			this.barButtonItem11 = new DevExpress.XtraBars.BarButtonItem();
			this.ribbonPage1 = new DevExpress.XtraBars.Ribbon.RibbonPage();
			this.ribbonPageGroup1 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
			this.ribbonPageGroup2 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
			this.ribbonPageGroup3 = new DevExpress.XtraBars.Ribbon.RibbonPageGroup();
			this.ribbonStatusBar = new DevExpress.XtraBars.Ribbon.RibbonStatusBar();
			this.documentManager1 = new DevExpress.XtraBars.Docking2010.DocumentManager(this.components);
			this.tabbedView1 = new DevExpress.XtraBars.Docking2010.Views.Tabbed.TabbedView(this.components);
			((System.ComponentModel.ISupportInitialize)(this.ribbon)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.documentManager1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.tabbedView1)).BeginInit();
			this.SuspendLayout();
			// 
			// ribbon
			// 
			this.ribbon.EmptyAreaImageOptions.ImagePadding = new System.Windows.Forms.Padding(63, 53, 63, 53);
			this.ribbon.ExpandCollapseItem.Id = 0;
			this.ribbon.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.ribbon.ExpandCollapseItem,
            this.ribbon.SearchEditItem,
            this.barButtonItem1,
            this.barButtonItem2,
            this.barButtonItem3,
            this.barButtonItem4,
            this.barButtonItem5,
            this.barButtonItem6,
            this.barButtonItem7,
            this.barButtonItem8,
            this.barButtonItem9,
            this.barButtonItem10,
            this.barButtonItem11});
			this.ribbon.Location = new System.Drawing.Point(0, 0);
			this.ribbon.Margin = new System.Windows.Forms.Padding(7, 5, 7, 5);
			this.ribbon.MaxItemId = 12;
			this.ribbon.Name = "ribbon";
			this.ribbon.OptionsMenuMinWidth = 692;
			this.ribbon.Pages.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPage[] {
            this.ribbonPage1});
			this.ribbon.Size = new System.Drawing.Size(1938, 158);
			this.ribbon.StatusBar = this.ribbonStatusBar;
			// 
			// barButtonItem1
			// 
			this.barButtonItem1.Caption = "Xuất file Excel quản lý sách";
			this.barButtonItem1.Id = 1;
			this.barButtonItem1.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("barButtonItem1.ImageOptions.Image")));
			this.barButtonItem1.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("barButtonItem1.ImageOptions.LargeImage")));
			this.barButtonItem1.Name = "barButtonItem1";
			this.barButtonItem1.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem1_ItemClick);
			// 
			// barButtonItem2
			// 
			this.barButtonItem2.Caption = "Xuất file Excel mượn / trả sách";
			this.barButtonItem2.Id = 2;
			this.barButtonItem2.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("barButtonItem2.ImageOptions.Image")));
			this.barButtonItem2.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("barButtonItem2.ImageOptions.LargeImage")));
			this.barButtonItem2.Name = "barButtonItem2";
			this.barButtonItem2.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem2_ItemClick);
			// 
			// barButtonItem3
			// 
			this.barButtonItem3.Caption = "Xuất file Excel tổng hợp";
			this.barButtonItem3.Id = 3;
			this.barButtonItem3.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("barButtonItem3.ImageOptions.Image")));
			this.barButtonItem3.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("barButtonItem3.ImageOptions.LargeImage")));
			this.barButtonItem3.Name = "barButtonItem3";
			this.barButtonItem3.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem3_ItemClick);
			// 
			// barButtonItem4
			// 
			this.barButtonItem4.Caption = "Quản lý sách";
			this.barButtonItem4.Id = 4;
			this.barButtonItem4.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("barButtonItem4.ImageOptions.Image")));
			this.barButtonItem4.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("barButtonItem4.ImageOptions.LargeImage")));
			this.barButtonItem4.Name = "barButtonItem4";
			this.barButtonItem4.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem4_ItemClick);
			// 
			// barButtonItem5
			// 
			this.barButtonItem5.Caption = "Mượn trả sách";
			this.barButtonItem5.Id = 5;
			this.barButtonItem5.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("barButtonItem5.ImageOptions.Image")));
			this.barButtonItem5.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("barButtonItem5.ImageOptions.LargeImage")));
			this.barButtonItem5.Name = "barButtonItem5";
			this.barButtonItem5.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem5_ItemClick);
			// 
			// barButtonItem6
			// 
			this.barButtonItem6.Caption = "Biểu đồ thống kê";
			this.barButtonItem6.Id = 6;
			this.barButtonItem6.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("barButtonItem6.ImageOptions.Image")));
			this.barButtonItem6.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("barButtonItem6.ImageOptions.LargeImage")));
			this.barButtonItem6.Name = "barButtonItem6";
			this.barButtonItem6.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem6_ItemClick);
			// 
			// barButtonItem7
			// 
			this.barButtonItem7.Caption = "barButtonItem7";
			this.barButtonItem7.Id = 7;
			this.barButtonItem7.Name = "barButtonItem7";
			this.barButtonItem7.Tag = "";
			// 
			// barButtonItem8
			// 
			this.barButtonItem8.Caption = "Xuất báo cáo quản lý sách";
			this.barButtonItem8.Id = 8;
			this.barButtonItem8.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("barButtonItem8.ImageOptions.SvgImage")));
			this.barButtonItem8.Name = "barButtonItem8";
			this.barButtonItem8.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem8_ItemClick);
			// 
			// barButtonItem9
			// 
			this.barButtonItem9.Caption = "Xuất báo cáo mượn / trả sách";
			this.barButtonItem9.Id = 9;
			this.barButtonItem9.ImageOptions.SvgImage = ((DevExpress.Utils.Svg.SvgImage)(resources.GetObject("barButtonItem9.ImageOptions.SvgImage")));
			this.barButtonItem9.Name = "barButtonItem9";
			this.barButtonItem9.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem9_ItemClick);
			// 
			// barButtonItem10
			// 
			this.barButtonItem10.Caption = "Thống kê số lượng sách đã mượn / trả nhiều nhất";
			this.barButtonItem10.Id = 10;
			this.barButtonItem10.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("barButtonItem10.ImageOptions.Image")));
			this.barButtonItem10.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("barButtonItem10.ImageOptions.LargeImage")));
			this.barButtonItem10.Name = "barButtonItem10";
			this.barButtonItem10.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem10_ItemClick);
			// 
			// barButtonItem11
			// 
			this.barButtonItem11.Caption = "Thống kê theo từng thông tin sách";
			this.barButtonItem11.Id = 11;
			this.barButtonItem11.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("barButtonItem11.ImageOptions.Image")));
			this.barButtonItem11.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("barButtonItem11.ImageOptions.LargeImage")));
			this.barButtonItem11.Name = "barButtonItem11";
			this.barButtonItem11.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem11_ItemClick);
			// 
			// ribbonPage1
			// 
			this.ribbonPage1.Groups.AddRange(new DevExpress.XtraBars.Ribbon.RibbonPageGroup[] {
            this.ribbonPageGroup1,
            this.ribbonPageGroup2,
            this.ribbonPageGroup3});
			this.ribbonPage1.Name = "ribbonPage1";
			this.ribbonPage1.Text = "Danh mục";
			// 
			// ribbonPageGroup1
			// 
			this.ribbonPageGroup1.ItemLinks.Add(this.barButtonItem1);
			this.ribbonPageGroup1.ItemLinks.Add(this.barButtonItem2);
			this.ribbonPageGroup1.ItemLinks.Add(this.barButtonItem3);
			this.ribbonPageGroup1.Name = "ribbonPageGroup1";
			this.ribbonPageGroup1.Text = "Xuất file Excel";
			// 
			// ribbonPageGroup2
			// 
			this.ribbonPageGroup2.ItemLinks.Add(this.barButtonItem4);
			this.ribbonPageGroup2.ItemLinks.Add(this.barButtonItem5);
			this.ribbonPageGroup2.ItemLinks.Add(this.barButtonItem10);
			this.ribbonPageGroup2.ItemLinks.Add(this.barButtonItem11);
			this.ribbonPageGroup2.ItemLinks.Add(this.barButtonItem6);
			this.ribbonPageGroup2.Name = "ribbonPageGroup2";
			this.ribbonPageGroup2.Text = "Thống kê thông tin";
			// 
			// ribbonPageGroup3
			// 
			this.ribbonPageGroup3.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("ribbonPageGroup3.ImageOptions.Image")));
			this.ribbonPageGroup3.ItemLinks.Add(this.barButtonItem8);
			this.ribbonPageGroup3.ItemLinks.Add(this.barButtonItem9);
			this.ribbonPageGroup3.Name = "ribbonPageGroup3";
			this.ribbonPageGroup3.Text = "Xuất file dạng báo cáo";
			// 
			// ribbonStatusBar
			// 
			this.ribbonStatusBar.Location = new System.Drawing.Point(0, 1075);
			this.ribbonStatusBar.Margin = new System.Windows.Forms.Padding(7, 5, 7, 5);
			this.ribbonStatusBar.Name = "ribbonStatusBar";
			this.ribbonStatusBar.Ribbon = this.ribbon;
			this.ribbonStatusBar.Size = new System.Drawing.Size(1938, 24);
			// 
			// documentManager1
			// 
			this.documentManager1.MdiParent = this;
			this.documentManager1.MenuManager = this.ribbon;
			this.documentManager1.View = this.tabbedView1;
			this.documentManager1.ViewCollection.AddRange(new DevExpress.XtraBars.Docking2010.Views.BaseView[] {
            this.tabbedView1});
			// 
			// bangthongke
			// 
			this.Appearance.ForeColor = System.Drawing.Color.White;
			this.Appearance.Options.UseFont = true;
			this.Appearance.Options.UseForeColor = true;
			this.AutoScaleDimensions = new System.Drawing.SizeF(12F, 23F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1938, 1099);
			this.Controls.Add(this.ribbonStatusBar);
			this.Controls.Add(this.ribbon);
			this.Font = new System.Drawing.Font("Tahoma", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.IconOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("bangthongke.IconOptions.LargeImage")));
			this.IsMdiContainer = true;
			this.Margin = new System.Windows.Forms.Padding(7, 5, 7, 5);
			this.MaximizeBox = false;
			this.Name = "bangthongke";
			this.Ribbon = this.ribbon;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.StatusBar = this.ribbonStatusBar;
			this.Text = "TRUNG TÂM THÔNG TIN THƯ VIỆN - TRƯỜNG ĐẠI HỌC HÀ TĨNH";
			((System.ComponentModel.ISupportInitialize)(this.ribbon)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.documentManager1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.tabbedView1)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private DevExpress.XtraBars.Ribbon.RibbonControl ribbon;
		private DevExpress.XtraBars.Ribbon.RibbonPage ribbonPage1;
		private DevExpress.XtraBars.Ribbon.RibbonPageGroup ribbonPageGroup1;
		private DevExpress.XtraBars.Ribbon.RibbonStatusBar ribbonStatusBar;
		private DevExpress.XtraBars.BarButtonItem barButtonItem1;
		private DevExpress.XtraBars.BarButtonItem barButtonItem2;
		private DevExpress.XtraBars.BarButtonItem barButtonItem3;
		private DevExpress.XtraBars.BarButtonItem barButtonItem4;
		private DevExpress.XtraBars.Ribbon.RibbonPageGroup ribbonPageGroup2;
		private DevExpress.XtraBars.BarButtonItem barButtonItem5;
		private DevExpress.XtraBars.BarButtonItem barButtonItem6;
		private DevExpress.XtraBars.Docking2010.DocumentManager documentManager1;
		private DevExpress.XtraBars.Docking2010.Views.Tabbed.TabbedView tabbedView1;
		private DevExpress.XtraBars.BarButtonItem barButtonItem7;
		private DevExpress.XtraBars.BarButtonItem barButtonItem8;
		private DevExpress.XtraBars.BarButtonItem barButtonItem9;
		private DevExpress.XtraBars.Ribbon.RibbonPageGroup ribbonPageGroup3;
		private DevExpress.XtraBars.BarButtonItem barButtonItem10;
		private DevExpress.XtraBars.BarButtonItem barButtonItem11;
	}
}