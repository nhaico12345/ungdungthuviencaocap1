namespace ungdungthuviencaocap
{
	partial class muontranhieunhat
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
			this.comboBox_thongkemuontra = new System.Windows.Forms.ComboBox();
			this.dataGridView_muontrasach = new System.Windows.Forms.DataGridView();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.dateTimePicker_chonngay = new System.Windows.Forms.DateTimePicker();
			this.radioDay = new System.Windows.Forms.RadioButton();
			this.radioMonth = new System.Windows.Forms.RadioButton();
			this.radioYear = new System.Windows.Forms.RadioButton();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.textBox_timkiem = new System.Windows.Forms.TextBox();
			this.Button_timKiem = new System.Windows.Forms.Button();
			this.Button_locngay = new System.Windows.Forms.Button();
			this.label5 = new System.Windows.Forms.Label();
			this.comboBox1 = new System.Windows.Forms.ComboBox();
			this.button_hientatcadulieu = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView_muontrasach)).BeginInit();
			this.SuspendLayout();
			// 
			// comboBox_thongkemuontra
			// 
			this.comboBox_thongkemuontra.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBox_thongkemuontra.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.comboBox_thongkemuontra.FormattingEnabled = true;
			this.comboBox_thongkemuontra.Items.AddRange(new object[] {
            "Thống kê mượn sách",
            "Thống kê trả sách"});
			this.comboBox_thongkemuontra.Location = new System.Drawing.Point(10, 187);
			this.comboBox_thongkemuontra.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.comboBox_thongkemuontra.Name = "comboBox_thongkemuontra";
			this.comboBox_thongkemuontra.Size = new System.Drawing.Size(247, 32);
			this.comboBox_thongkemuontra.TabIndex = 0;
			this.comboBox_thongkemuontra.SelectedIndexChanged += new System.EventHandler(this.comboBox_thongkemuontra_SelectedIndexChanged);
			// 
			// dataGridView_muontrasach
			// 
			this.dataGridView_muontrasach.AllowUserToAddRows = false;
			this.dataGridView_muontrasach.AllowUserToDeleteRows = false;
			this.dataGridView_muontrasach.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
			this.dataGridView_muontrasach.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
			this.dataGridView_muontrasach.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridView_muontrasach.Dock = System.Windows.Forms.DockStyle.Right;
			this.dataGridView_muontrasach.Location = new System.Drawing.Point(269, 0);
			this.dataGridView_muontrasach.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.dataGridView_muontrasach.Name = "dataGridView_muontrasach";
			this.dataGridView_muontrasach.ReadOnly = true;
			this.dataGridView_muontrasach.RowHeadersWidth = 51;
			this.dataGridView_muontrasach.Size = new System.Drawing.Size(1211, 816);
			this.dataGridView_muontrasach.TabIndex = 1;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(37, 160);
			this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(189, 24);
			this.label1.TabIndex = 2;
			this.label1.Text = "Chọn bảng thống kê";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label2.Location = new System.Drawing.Point(8, 364);
			this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(222, 24);
			this.label2.TabIndex = 3;
			this.label2.Text = "Chọn ngày/ tháng/ năm";
			// 
			// dateTimePicker_chonngay
			// 
			this.dateTimePicker_chonngay.Cursor = System.Windows.Forms.Cursors.Hand;
			this.dateTimePicker_chonngay.Font = new System.Drawing.Font("Tahoma", 10.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.dateTimePicker_chonngay.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
			this.dateTimePicker_chonngay.Location = new System.Drawing.Point(41, 405);
			this.dateTimePicker_chonngay.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
			this.dateTimePicker_chonngay.Name = "dateTimePicker_chonngay";
			this.dateTimePicker_chonngay.Size = new System.Drawing.Size(160, 29);
			this.dateTimePicker_chonngay.TabIndex = 4;
			this.dateTimePicker_chonngay.ValueChanged += new System.EventHandler(this.dateTimePicker_chonngay_ValueChanged);
			// 
			// radioDay
			// 
			this.radioDay.AutoSize = true;
			this.radioDay.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.radioDay.Location = new System.Drawing.Point(55, 262);
			this.radioDay.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
			this.radioDay.Name = "radioDay";
			this.radioDay.Size = new System.Drawing.Size(134, 25);
			this.radioDay.TabIndex = 5;
			this.radioDay.TabStop = true;
			this.radioDay.Text = "Lọc theo ngày";
			this.radioDay.UseVisualStyleBackColor = true;
			this.radioDay.CheckedChanged += new System.EventHandler(this.DateFilter_CheckedChanged);
			// 
			// radioMonth
			// 
			this.radioMonth.AutoSize = true;
			this.radioMonth.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.radioMonth.Location = new System.Drawing.Point(55, 293);
			this.radioMonth.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
			this.radioMonth.Name = "radioMonth";
			this.radioMonth.Size = new System.Drawing.Size(141, 25);
			this.radioMonth.TabIndex = 6;
			this.radioMonth.TabStop = true;
			this.radioMonth.Text = "Lọc theo tháng";
			this.radioMonth.UseVisualStyleBackColor = true;
			this.radioMonth.CheckedChanged += new System.EventHandler(this.DateFilter_CheckedChanged);
			// 
			// radioYear
			// 
			this.radioYear.AutoSize = true;
			this.radioYear.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.radioYear.Location = new System.Drawing.Point(55, 324);
			this.radioYear.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
			this.radioYear.Name = "radioYear";
			this.radioYear.Size = new System.Drawing.Size(131, 25);
			this.radioYear.TabIndex = 7;
			this.radioYear.TabStop = true;
			this.radioYear.Text = "Lọc theo năm";
			this.radioYear.UseVisualStyleBackColor = true;
			this.radioYear.CheckedChanged += new System.EventHandler(this.DateFilter_CheckedChanged);
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label3.Location = new System.Drawing.Point(37, 236);
			this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(180, 24);
			this.label3.TabIndex = 8;
			this.label3.Text = "Chọn định dạng lọc";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label4.Location = new System.Drawing.Point(16, 489);
			this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(212, 24);
			this.label4.TabIndex = 9;
			this.label4.Text = "Tìm kiếm mã sinh viên";
			// 
			// textBox_timkiem
			// 
			this.textBox_timkiem.Location = new System.Drawing.Point(21, 514);
			this.textBox_timkiem.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
			this.textBox_timkiem.Name = "textBox_timkiem";
			this.textBox_timkiem.Size = new System.Drawing.Size(204, 23);
			this.textBox_timkiem.TabIndex = 10;
			this.textBox_timkiem.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox_timkiem_KeyDown);
			// 
			// Button_timKiem
			// 
			this.Button_timKiem.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Button_timKiem.Location = new System.Drawing.Point(72, 558);
			this.Button_timKiem.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
			this.Button_timKiem.Name = "Button_timKiem";
			this.Button_timKiem.Size = new System.Drawing.Size(88, 34);
			this.Button_timKiem.TabIndex = 11;
			this.Button_timKiem.Text = "Tìm";
			this.Button_timKiem.UseVisualStyleBackColor = true;
			this.Button_timKiem.Click += new System.EventHandler(this.Button_timKiem_Click);
			// 
			// Button_locngay
			// 
			this.Button_locngay.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.Button_locngay.Location = new System.Drawing.Point(72, 441);
			this.Button_locngay.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
			this.Button_locngay.Name = "Button_locngay";
			this.Button_locngay.Size = new System.Drawing.Size(88, 34);
			this.Button_locngay.TabIndex = 12;
			this.Button_locngay.Text = "Lọc";
			this.Button_locngay.UseVisualStyleBackColor = true;
			this.Button_locngay.Click += new System.EventHandler(this.Button_locngay_Click);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label5.Location = new System.Drawing.Point(70, 81);
			this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(116, 24);
			this.label5.TabIndex = 14;
			this.label5.Text = "Chế độ xem";
			// 
			// comboBox1
			// 
			this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBox1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.comboBox1.FormattingEnabled = true;
			this.comboBox1.Items.AddRange(new object[] {
            "Chế độ xem tổng hợp",
            "Chế độ xem chi tiết"});
			this.comboBox1.Location = new System.Drawing.Point(12, 108);
			this.comboBox1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.comboBox1.Name = "comboBox1";
			this.comboBox1.Size = new System.Drawing.Size(247, 32);
			this.comboBox1.TabIndex = 13;
			this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
			// 
			// button_hientatcadulieu
			// 
			this.button_hientatcadulieu.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button_hientatcadulieu.Location = new System.Drawing.Point(39, 31);
			this.button_hientatcadulieu.Margin = new System.Windows.Forms.Padding(4, 2, 4, 2);
			this.button_hientatcadulieu.Name = "button_hientatcadulieu";
			this.button_hientatcadulieu.Size = new System.Drawing.Size(178, 34);
			this.button_hientatcadulieu.TabIndex = 15;
			this.button_hientatcadulieu.Text = "Hiện tất cả dữ liệu";
			this.button_hientatcadulieu.UseVisualStyleBackColor = true;
			this.button_hientatcadulieu.Click += new System.EventHandler(this.button_hientatcadulieu_Click);
			// 
			// muontranhieunhat
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1480, 816);
			this.Controls.Add(this.button_hientatcadulieu);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.comboBox1);
			this.Controls.Add(this.Button_locngay);
			this.Controls.Add(this.Button_timKiem);
			this.Controls.Add(this.textBox_timkiem);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.radioYear);
			this.Controls.Add(this.radioMonth);
			this.Controls.Add(this.radioDay);
			this.Controls.Add(this.dateTimePicker_chonngay);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.dataGridView_muontrasach);
			this.Controls.Add(this.comboBox_thongkemuontra);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.Name = "muontranhieunhat";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Thống kê mượn trả";
			((System.ComponentModel.ISupportInitialize)(this.dataGridView_muontrasach)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.ComboBox comboBox_thongkemuontra;
		private System.Windows.Forms.DataGridView dataGridView_muontrasach;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.DateTimePicker dateTimePicker_chonngay;
		private System.Windows.Forms.RadioButton radioDay;
		private System.Windows.Forms.RadioButton radioMonth;
		private System.Windows.Forms.RadioButton radioYear;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox textBox_timkiem;
		private System.Windows.Forms.Button Button_timKiem;
		private System.Windows.Forms.Button Button_locngay;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.ComboBox comboBox1;
		private System.Windows.Forms.Button button_hientatcadulieu;
	}
}