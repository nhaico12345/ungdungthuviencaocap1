namespace ungdungthuviencaocap
{
	partial class nhaxuatban
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
			this.label10 = new System.Windows.Forms.Label();
			this.button_loc = new System.Windows.Forms.Button();
			this.dateTimePicker_ngaythanhlap = new System.Windows.Forms.DateTimePicker();
			this.button_lammoi = new System.Windows.Forms.Button();
			this.button_capnhat = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.textBox_tennhaxuatban = new System.Windows.Forms.TextBox();
			this.dataGridView_nhaxb = new System.Windows.Forms.DataGridView();
			this.panel3 = new System.Windows.Forms.Panel();
			this.dataGridView_tacpham = new System.Windows.Forms.DataGridView();
			this.label11 = new System.Windows.Forms.Label();
			this.textBox_trangthai = new System.Windows.Forms.TextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.label9 = new System.Windows.Forms.Label();
			this.panel2 = new System.Windows.Forms.Panel();
			this.textBox_email = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.textBox_sohieuxuatban = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.textBox_diachi = new System.Windows.Forms.TextBox();
			this.label7 = new System.Windows.Forms.Label();
			this.label8 = new System.Windows.Forms.Label();
			this.panel1 = new System.Windows.Forms.Panel();
			this.textBox_websiste = new System.Windows.Forms.TextBox();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView_nhaxb)).BeginInit();
			this.panel3.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView_tacpham)).BeginInit();
			this.panel2.SuspendLayout();
			this.panel1.SuspendLayout();
			this.SuspendLayout();
			// 
			// label10
			// 
			this.label10.AutoSize = true;
			this.label10.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label10.Location = new System.Drawing.Point(246, 9);
			this.label10.Name = "label10";
			this.label10.Size = new System.Drawing.Size(179, 19);
			this.label10.TabIndex = 0;
			this.label10.Text = "Danh sách nhà xuất bản";
			// 
			// button_loc
			// 
			this.button_loc.Location = new System.Drawing.Point(116, 287);
			this.button_loc.Name = "button_loc";
			this.button_loc.Size = new System.Drawing.Size(86, 23);
			this.button_loc.TabIndex = 26;
			this.button_loc.Text = "Lọc";
			this.button_loc.UseVisualStyleBackColor = true;
			// 
			// dateTimePicker_ngaythanhlap
			// 
			this.dateTimePicker_ngaythanhlap.Format = System.Windows.Forms.DateTimePickerFormat.Custom; //Sửa ở đây
			this.dateTimePicker_ngaythanhlap.Location = new System.Drawing.Point(116, 75);
			this.dateTimePicker_ngaythanhlap.Name = "dateTimePicker_ngaythanhlap";
			this.dateTimePicker_ngaythanhlap.Size = new System.Drawing.Size(355, 21);
			this.dateTimePicker_ngaythanhlap.TabIndex = 25;
			// 
			// button_lammoi
			// 
			this.button_lammoi.Location = new System.Drawing.Point(385, 287);
			this.button_lammoi.Name = "button_lammoi";
			this.button_lammoi.Size = new System.Drawing.Size(86, 23);
			this.button_lammoi.TabIndex = 22;
			this.button_lammoi.Text = "Làm mới";
			this.button_lammoi.UseVisualStyleBackColor = true;
			this.button_lammoi.Click += new System.EventHandler(this.button_lammoi_Click); // Thêm sự kiện
																						   // 
																						   // button_capnhat
																						   // 
			this.button_capnhat.Location = new System.Drawing.Point(238, 287);
			this.button_capnhat.Name = "button_capnhat";
			this.button_capnhat.Size = new System.Drawing.Size(117, 23);
			this.button_capnhat.TabIndex = 20;
			this.button_capnhat.Text = "Cập nhật thông tin";
			this.button_capnhat.UseVisualStyleBackColor = true;
			// this.button_capnhat.Click += new System.EventHandler(this.button_capnhat_Click); // Đã đăng ký trong constructor của form
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label2.Location = new System.Drawing.Point(3, 41);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(91, 18);
			this.label2.TabIndex = 1;
			this.label2.Text = "Tên nhà XB:";
			// 
			// textBox_tennhaxuatban
			// 
			this.textBox_tennhaxuatban.Location = new System.Drawing.Point(116, 38);
			this.textBox_tennhaxuatban.Name = "textBox_tennhaxuatban";
			this.textBox_tennhaxuatban.Size = new System.Drawing.Size(355, 21);
			this.textBox_tennhaxuatban.TabIndex = 4;
			// 
			// dataGridView_nhaxb
			// 
			this.dataGridView_nhaxb.AllowUserToAddRows = false;
			this.dataGridView_nhaxb.AllowUserToDeleteRows = false;
			this.dataGridView_nhaxb.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
			this.dataGridView_nhaxb.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
			this.dataGridView_nhaxb.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridView_nhaxb.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.dataGridView_nhaxb.Location = new System.Drawing.Point(0, 31);
			this.dataGridView_nhaxb.Name = "dataGridView_nhaxb";
			this.dataGridView_nhaxb.ReadOnly = true;
			this.dataGridView_nhaxb.Size = new System.Drawing.Size(674, 290);
			this.dataGridView_nhaxb.TabIndex = 1;
			// this.dataGridView_nhaxb.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_nhaxb_CellClick); // Đã đăng ký trong constructor
			// 
			// panel3
			// 
			this.panel3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
			this.panel3.Controls.Add(this.dataGridView_tacpham);
			this.panel3.Controls.Add(this.label11);
			this.panel3.Location = new System.Drawing.Point(-1, 327);
			this.panel3.Name = "panel3";
			this.panel3.Size = new System.Drawing.Size(1160, 350);
			this.panel3.TabIndex = 5;
			// 
			// dataGridView_tacpham
			// 
			this.dataGridView_tacpham.AllowUserToAddRows = false;
			this.dataGridView_tacpham.AllowUserToDeleteRows = false;
			this.dataGridView_tacpham.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
			this.dataGridView_tacpham.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
			this.dataGridView_tacpham.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridView_tacpham.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.dataGridView_tacpham.Location = new System.Drawing.Point(0, 31);
			this.dataGridView_tacpham.Name = "dataGridView_tacpham";
			this.dataGridView_tacpham.ReadOnly = true;
			this.dataGridView_tacpham.Size = new System.Drawing.Size(1160, 319);
			this.dataGridView_tacpham.TabIndex = 1;
			// 
			// label11
			// 
			this.label11.AutoSize = true;
			this.label11.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label11.Location = new System.Drawing.Point(551, 9);
			this.label11.Name = "label11";
			this.label11.Size = new System.Drawing.Size(152, 19);
			this.label11.TabIndex = 0;
			this.label11.Text = "Danh sách tác phẩm";
			// 
			// textBox_trangthai
			// 
			this.textBox_trangthai.Location = new System.Drawing.Point(116, 217);
			this.textBox_trangthai.Name = "textBox_trangthai";
			this.textBox_trangthai.Size = new System.Drawing.Size(355, 21);
			this.textBox_trangthai.TabIndex = 16;
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label3.Location = new System.Drawing.Point(3, 78);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(80, 18);
			this.label3.TabIndex = 5;
			this.label3.Text = "Ngày Tlập:";
			// 
			// label9
			// 
			this.label9.AutoSize = true;
			this.label9.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label9.Location = new System.Drawing.Point(3, 220);
			this.label9.Name = "label9";
			this.label9.Size = new System.Drawing.Size(80, 18);
			this.label9.TabIndex = 15;
			this.label9.Text = "Trạng thái:";
			// 
			// panel2
			// 
			this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
			this.panel2.Controls.Add(this.dataGridView_nhaxb);
			this.panel2.Controls.Add(this.label10);
			this.panel2.Location = new System.Drawing.Point(-1, 0);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(674, 321);
			this.panel2.TabIndex = 4;
			// 
			// textBox_email
			// 
			this.textBox_email.Location = new System.Drawing.Point(116, 181);
			this.textBox_email.Name = "textBox_email";
			this.textBox_email.Size = new System.Drawing.Size(355, 21);
			this.textBox_email.TabIndex = 14;
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label4.Location = new System.Drawing.Point(3, 149);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(83, 18);
			this.label4.TabIndex = 7;
			this.label4.Text = "Số hiệu XB:";
			// 
			// textBox_sohieuxuatban
			// 
			this.textBox_sohieuxuatban.Location = new System.Drawing.Point(116, 146);
			this.textBox_sohieuxuatban.Name = "textBox_sohieuxuatban";
			this.textBox_sohieuxuatban.Size = new System.Drawing.Size(355, 21);
			this.textBox_sohieuxuatban.TabIndex = 8;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(187, 9);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(175, 19);
			this.label1.TabIndex = 0;
			this.label1.Text = "Thông tin nhà xuất bản";
			// 
			// textBox_diachi
			// 
			this.textBox_diachi.Location = new System.Drawing.Point(116, 109);
			this.textBox_diachi.Name = "textBox_diachi";
			this.textBox_diachi.Size = new System.Drawing.Size(355, 21);
			this.textBox_diachi.TabIndex = 12;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label7.Location = new System.Drawing.Point(3, 112);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(55, 18);
			this.label7.TabIndex = 11;
			this.label7.Text = "Địa chỉ:";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label8.Location = new System.Drawing.Point(3, 184);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(46, 18);
			this.label8.TabIndex = 13;
			this.label8.Text = "Email:";
			// 
			// panel1
			// 
			this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
			this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
			this.panel1.Controls.Add(this.textBox_websiste);
			this.panel1.Controls.Add(this.label5);
			this.panel1.Controls.Add(this.button_loc);
			this.panel1.Controls.Add(this.dateTimePicker_ngaythanhlap);
			this.panel1.Controls.Add(this.button_lammoi);
			this.panel1.Controls.Add(this.button_capnhat);
			this.panel1.Controls.Add(this.label2);
			this.panel1.Controls.Add(this.textBox_tennhaxuatban);
			this.panel1.Controls.Add(this.textBox_trangthai);
			this.panel1.Controls.Add(this.label3);
			this.panel1.Controls.Add(this.label9);
			this.panel1.Controls.Add(this.textBox_email);
			this.panel1.Controls.Add(this.label4);
			this.panel1.Controls.Add(this.label8);
			this.panel1.Controls.Add(this.textBox_sohieuxuatban);
			this.panel1.Controls.Add(this.label6); // Giữ lại label6 hoặc label1, tùy theo ý bạn muốn hiển thị tiêu đề nào
			this.panel1.Controls.Add(this.label1); // Hoặc label1
			this.panel1.Controls.Add(this.textBox_diachi);
			this.panel1.Controls.Add(this.label7);
			this.panel1.Location = new System.Drawing.Point(676, 0);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(483, 321);
			this.panel1.TabIndex = 3;
			// 
			// textBox_websiste
			// 
			this.textBox_websiste.Location = new System.Drawing.Point(116, 251);
			this.textBox_websiste.Name = "textBox_websiste";
			this.textBox_websiste.Size = new System.Drawing.Size(355, 21);
			this.textBox_websiste.TabIndex = 28;
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label5.Location = new System.Drawing.Point(3, 254);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(65, 18);
			this.label5.TabIndex = 27;
			this.label5.Text = "Website:";
			// 
			// label6 
			// Xóa hoặc comment out label6 nếu bạn chỉ muốn dùng label1 cho tiêu đề panel1
			// this.label6.AutoSize = true;
			// this.label6.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			// this.label6.Location = new System.Drawing.Point(187, 9);
			// this.label6.Name = "label6";
			// this.label6.Size = new System.Drawing.Size(175, 19);
			// this.label6.TabIndex = 0;
			// this.label6.Text = "Thông tin nhà xuất bản"; 
			// 
			// nhaxuatban
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1159, 677);
			this.Controls.Add(this.panel3);
			this.Controls.Add(this.panel2);
			this.Controls.Add(this.panel1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			this.Name = "nhaxuatban";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Quản lý Nhà Xuất Bản"; // Sửa tiêu đề form
			this.Load += new System.EventHandler(this.nhaxuatban_Load);
			((System.ComponentModel.ISupportInitialize)(this.dataGridView_nhaxb)).EndInit();
			this.panel3.ResumeLayout(false);
			this.panel3.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView_tacpham)).EndInit();
			this.panel2.ResumeLayout(false);
			this.panel2.PerformLayout();
			this.panel1.ResumeLayout(false);
			this.panel1.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Label label10;
		private System.Windows.Forms.Button button_loc;
		private System.Windows.Forms.DateTimePicker dateTimePicker_ngaythanhlap;
		private System.Windows.Forms.Button button_lammoi;
		private System.Windows.Forms.Button button_capnhat;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox textBox_tennhaxuatban;
		private System.Windows.Forms.DataGridView dataGridView_nhaxb;
		private System.Windows.Forms.Panel panel3;
		private System.Windows.Forms.DataGridView dataGridView_tacpham;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.TextBox textBox_trangthai;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Panel panel2;
		private System.Windows.Forms.TextBox textBox_email;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox textBox_sohieuxuatban;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox textBox_diachi;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label label8;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.TextBox textBox_websiste;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6; // Đã comment out hoặc xóa nếu không cần
	}
}
