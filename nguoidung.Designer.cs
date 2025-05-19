namespace ungdungthuviencaocap
{
	partial class nguoidung
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(nguoidung));
			this.fluentDesignFormContainer1 = new DevExpress.XtraBars.FluentDesignSystem.FluentDesignFormContainer();
			this.label1 = new System.Windows.Forms.Label();
			this.accordionControl1 = new DevExpress.XtraBars.Navigation.AccordionControl();
			this.mnTrangchu = new DevExpress.XtraBars.Navigation.AccordionControlElement();
			this.mnChucNang = new DevExpress.XtraBars.Navigation.AccordionControlElement();
			this.mnMuonTraSach = new DevExpress.XtraBars.Navigation.AccordionControlElement();
			this.mnTonghop = new DevExpress.XtraBars.Navigation.AccordionControlElement();
			this.mnHelp = new DevExpress.XtraBars.Navigation.AccordionControlElement();
			this.mnCaiDat = new DevExpress.XtraBars.Navigation.AccordionControlElement();
			this.mnThongTin = new DevExpress.XtraBars.Navigation.AccordionControlElement();
			this.mnDangXuat = new DevExpress.XtraBars.Navigation.AccordionControlElement();
			this.fluentDesignFormControl1 = new DevExpress.XtraBars.FluentDesignSystem.FluentDesignFormControl();
			this.lblTieude = new DevExpress.XtraBars.BarStaticItem();
			this.fluentFormDefaultManager1 = new DevExpress.XtraBars.FluentDesignSystem.FluentFormDefaultManager(this.components);
			this.mnThongTinTaiKhoan = new DevExpress.XtraBars.Navigation.AccordionControlElement();
			this.fluentDesignFormContainer1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.accordionControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.fluentDesignFormControl1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.fluentFormDefaultManager1)).BeginInit();
			this.SuspendLayout();
			// 
			// fluentDesignFormContainer1
			// 
			this.fluentDesignFormContainer1.Controls.Add(this.label1);
			this.fluentDesignFormContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.fluentDesignFormContainer1.Location = new System.Drawing.Point(223, 31);
			this.fluentDesignFormContainer1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.fluentDesignFormContainer1.Name = "fluentDesignFormContainer1";
			this.fluentDesignFormContainer1.Size = new System.Drawing.Size(1141, 618);
			this.fluentDesignFormContainer1.TabIndex = 0;
			// 
			// label1
			// 
			this.label1.Anchor = System.Windows.Forms.AnchorStyles.None;
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 36F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(357, 264);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(475, 116);
			this.label1.TabIndex = 1;
			this.label1.Text = "     Welcome To\r\nỨng dụng thư viện";
			// 
			// accordionControl1
			// 
			this.accordionControl1.Dock = System.Windows.Forms.DockStyle.Left;
			this.accordionControl1.Elements.AddRange(new DevExpress.XtraBars.Navigation.AccordionControlElement[] {
            this.mnTrangchu,
            this.mnChucNang,
            this.mnCaiDat,
            this.mnThongTinTaiKhoan,
            this.mnThongTin,
            this.mnDangXuat});
			this.accordionControl1.Location = new System.Drawing.Point(0, 31);
			this.accordionControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.accordionControl1.Name = "accordionControl1";
			this.accordionControl1.ScrollBarMode = DevExpress.XtraBars.Navigation.ScrollBarMode.Touch;
			this.accordionControl1.Size = new System.Drawing.Size(223, 618);
			this.accordionControl1.TabIndex = 1;
			this.accordionControl1.ViewType = DevExpress.XtraBars.Navigation.AccordionControlViewType.HamburgerMenu;
			// 
			// mnTrangchu
			// 
			this.mnTrangchu.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("mnTrangchu.ImageOptions.Image")));
			this.mnTrangchu.Name = "mnTrangchu";
			this.mnTrangchu.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
			this.mnTrangchu.Text = "Trang chủ";
			this.mnTrangchu.Click += new System.EventHandler(this.mnTrangchu_Click);
			// 
			// mnChucNang
			// 
			this.mnChucNang.Elements.AddRange(new DevExpress.XtraBars.Navigation.AccordionControlElement[] {
            this.mnMuonTraSach,
            this.mnTonghop,
            this.mnHelp});
			this.mnChucNang.Expanded = true;
			this.mnChucNang.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("mnChucNang.ImageOptions.Image")));
			this.mnChucNang.Name = "mnChucNang";
			this.mnChucNang.Text = "Chức năng";
			// 
			// mnMuonTraSach
			// 
			this.mnMuonTraSach.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("mnMuonTraSach.ImageOptions.Image")));
			this.mnMuonTraSach.Name = "mnMuonTraSach";
			this.mnMuonTraSach.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
			this.mnMuonTraSach.Text = "Hẹn lịch mượn sách";
			this.mnMuonTraSach.Click += new System.EventHandler(this.mnMuonTraSach_Click);
			// 
			// mnTonghop
			// 
			this.mnTonghop.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("mnTonghop.ImageOptions.Image")));
			this.mnTonghop.Name = "mnTonghop";
			this.mnTonghop.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
			this.mnTonghop.Text = "Tổng hợp sách thư viện";
			this.mnTonghop.Click += new System.EventHandler(this.accordionControlElement1_Click);
			// 
			// mnHelp
			// 
			this.mnHelp.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("mnHelp.ImageOptions.Image")));
			this.mnHelp.Name = "mnHelp";
			this.mnHelp.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
			this.mnHelp.Text = "Giải đáp thắc mắc";
			this.mnHelp.Click += new System.EventHandler(this.mnHelp_Click);
			// 
			// mnCaiDat
			// 
			this.mnCaiDat.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("mnCaiDat.ImageOptions.Image")));
			this.mnCaiDat.Name = "mnCaiDat";
			this.mnCaiDat.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
			this.mnCaiDat.Text = "Cài đặt";
			// 
			// mnThongTin
			// 
			this.mnThongTin.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("mnThongTin.ImageOptions.Image")));
			this.mnThongTin.Name = "mnThongTin";
			this.mnThongTin.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
			this.mnThongTin.Text = "Thông tin liên hệ";
			this.mnThongTin.Click += new System.EventHandler(this.mnThongTin_Click);
			// 
			// mnDangXuat
			// 
			this.mnDangXuat.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("mnDangXuat.ImageOptions.Image")));
			this.mnDangXuat.Name = "mnDangXuat";
			this.mnDangXuat.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
			this.mnDangXuat.Text = "Đăng xuẩt";
			this.mnDangXuat.Click += new System.EventHandler(this.mnDangXuat_Click);
			// 
			// fluentDesignFormControl1
			// 
			this.fluentDesignFormControl1.FluentDesignForm = this;
			this.fluentDesignFormControl1.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.lblTieude});
			this.fluentDesignFormControl1.Location = new System.Drawing.Point(0, 0);
			this.fluentDesignFormControl1.Manager = this.fluentFormDefaultManager1;
			this.fluentDesignFormControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.fluentDesignFormControl1.Name = "fluentDesignFormControl1";
			this.fluentDesignFormControl1.Size = new System.Drawing.Size(1364, 31);
			this.fluentDesignFormControl1.TabIndex = 2;
			this.fluentDesignFormControl1.TabStop = false;
			this.fluentDesignFormControl1.TitleItemLinks.Add(this.lblTieude);
			// 
			// lblTieude
			// 
			this.lblTieude.Caption = "Demo thư viện - Tác giả: Trần Đình Đăng";
			this.lblTieude.Id = 0;
			this.lblTieude.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("lblTieude.ImageOptions.Image")));
			this.lblTieude.ImageOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("lblTieude.ImageOptions.LargeImage")));
			this.lblTieude.ItemAppearance.Normal.Font = new System.Drawing.Font("Tahoma", 7.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.lblTieude.ItemAppearance.Normal.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(0)))), ((int)(((byte)(192)))));
			this.lblTieude.ItemAppearance.Normal.Options.UseFont = true;
			this.lblTieude.ItemAppearance.Normal.Options.UseForeColor = true;
			this.lblTieude.Name = "lblTieude";
			this.lblTieude.PaintStyle = DevExpress.XtraBars.BarItemPaintStyle.CaptionGlyph;
			// 
			// fluentFormDefaultManager1
			// 
			this.fluentFormDefaultManager1.Form = this;
			this.fluentFormDefaultManager1.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.lblTieude});
			this.fluentFormDefaultManager1.MaxItemId = 1;
			// 
			// mnThongTinTaiKhoan
			// 
			this.mnThongTinTaiKhoan.ImageOptions.Image = ((System.Drawing.Image)(resources.GetObject("accordionControlElement1.ImageOptions.Image")));
			this.mnThongTinTaiKhoan.Name = "mnThongTinTaiKhoan";
			this.mnThongTinTaiKhoan.Style = DevExpress.XtraBars.Navigation.ElementStyle.Item;
			this.mnThongTinTaiKhoan.Text = "Thông tin tài khoản";
			this.mnThongTinTaiKhoan.Click += new System.EventHandler(this.accordionControlElement1_Click_1);
			// 
			// nguoidung
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1364, 649);
			this.ControlContainer = this.fluentDesignFormContainer1;
			this.Controls.Add(this.fluentDesignFormContainer1);
			this.Controls.Add(this.accordionControl1);
			this.Controls.Add(this.fluentDesignFormControl1);
			this.FluentDesignFormControl = this.fluentDesignFormControl1;
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.IconOptions.Image = global::ungdungthuviencaocap.Properties.Resources.icon1;
			this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.MaximizeBox = false;
			this.Name = "nguoidung";
			this.NavigationControl = this.accordionControl1;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Thư viện";
			this.fluentDesignFormContainer1.ResumeLayout(false);
			this.fluentDesignFormContainer1.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.accordionControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.fluentDesignFormControl1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.fluentFormDefaultManager1)).EndInit();
			this.ResumeLayout(false);

		}

		#endregion
		private DevExpress.XtraBars.FluentDesignSystem.FluentDesignFormContainer fluentDesignFormContainer1;
		private DevExpress.XtraBars.Navigation.AccordionControl accordionControl1;
		private DevExpress.XtraBars.FluentDesignSystem.FluentDesignFormControl fluentDesignFormControl1;
		private DevExpress.XtraBars.Navigation.AccordionControlElement mnTrangchu;
		private DevExpress.XtraBars.FluentDesignSystem.FluentFormDefaultManager fluentFormDefaultManager1;
		private System.Windows.Forms.Label label1;
		private DevExpress.XtraBars.Navigation.AccordionControlElement mnChucNang;
		private DevExpress.XtraBars.Navigation.AccordionControlElement mnMuonTraSach;
		private DevExpress.XtraBars.Navigation.AccordionControlElement mnHelp;
		private DevExpress.XtraBars.Navigation.AccordionControlElement mnCaiDat;
		private DevExpress.XtraBars.Navigation.AccordionControlElement mnThongTin;
		private DevExpress.XtraBars.Navigation.AccordionControlElement mnDangXuat;
		private DevExpress.XtraBars.BarStaticItem lblTieude;
		private DevExpress.XtraBars.Navigation.AccordionControlElement mnTonghop;
		private DevExpress.XtraBars.Navigation.AccordionControlElement mnThongTinTaiKhoan;
	}
}