namespace ungdungthuviencaocap
{
	partial class khoamotaikhoan
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
			this.dataGridView_taikhoan = new System.Windows.Forms.DataGridView();
			this.button_khoataikhoan = new System.Windows.Forms.Button();
			this.button_motaikhoan = new System.Windows.Forms.Button();
			this.textBox_timtendangnhap = new System.Windows.Forms.TextBox();
			this.textBox_timmasinhvien = new System.Windows.Forms.TextBox();
			this.button3 = new System.Windows.Forms.Button();
			this.button4 = new System.Windows.Forms.Button();
			this.button1 = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView_taikhoan)).BeginInit();
			this.SuspendLayout();
			// 
			// dataGridView_taikhoan
			// 
			this.dataGridView_taikhoan.AllowUserToAddRows = false;
			this.dataGridView_taikhoan.AllowUserToDeleteRows = false;
			this.dataGridView_taikhoan.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
			this.dataGridView_taikhoan.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
			this.dataGridView_taikhoan.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridView_taikhoan.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.dataGridView_taikhoan.Location = new System.Drawing.Point(0, 170);
			this.dataGridView_taikhoan.Name = "dataGridView_taikhoan";
			this.dataGridView_taikhoan.ReadOnly = true;
			this.dataGridView_taikhoan.Size = new System.Drawing.Size(925, 215);
			this.dataGridView_taikhoan.TabIndex = 21;
			// 
			// button_khoataikhoan
			// 
			this.button_khoataikhoan.Location = new System.Drawing.Point(287, 127);
			this.button_khoataikhoan.Name = "button_khoataikhoan";
			this.button_khoataikhoan.Size = new System.Drawing.Size(94, 26);
			this.button_khoataikhoan.TabIndex = 22;
			this.button_khoataikhoan.Text = "Khóa tài khoản";
			this.button_khoataikhoan.UseVisualStyleBackColor = true;
			this.button_khoataikhoan.Click += new System.EventHandler(this.button_khoataikhoan_Click);
			// 
			// button_motaikhoan
			// 
			this.button_motaikhoan.Location = new System.Drawing.Point(554, 127);
			this.button_motaikhoan.Name = "button_motaikhoan";
			this.button_motaikhoan.Size = new System.Drawing.Size(94, 26);
			this.button_motaikhoan.TabIndex = 23;
			this.button_motaikhoan.Text = "Mở tài khoản";
			this.button_motaikhoan.UseVisualStyleBackColor = true;
			this.button_motaikhoan.Click += new System.EventHandler(this.button_motaikhoan_Click);
			// 
			// textBox_timtendangnhap
			// 
			this.textBox_timtendangnhap.Location = new System.Drawing.Point(455, 16);
			this.textBox_timtendangnhap.Name = "textBox_timtendangnhap";
			this.textBox_timtendangnhap.Size = new System.Drawing.Size(212, 21);
			this.textBox_timtendangnhap.TabIndex = 24;
			this.textBox_timtendangnhap.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox_timtendangnhap_KeyDown);
			// 
			// textBox_timmasinhvien
			// 
			this.textBox_timmasinhvien.Location = new System.Drawing.Point(455, 63);
			this.textBox_timmasinhvien.Name = "textBox_timmasinhvien";
			this.textBox_timmasinhvien.Size = new System.Drawing.Size(212, 21);
			this.textBox_timmasinhvien.TabIndex = 25;
			this.textBox_timmasinhvien.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox_timmasinhvien_KeyDown);
			// 
			// button3
			// 
			this.button3.Location = new System.Drawing.Point(287, 12);
			this.button3.Name = "button3";
			this.button3.Size = new System.Drawing.Size(139, 26);
			this.button3.TabIndex = 26;
			this.button3.Text = "Tìm theo tên đăng nhập";
			this.button3.UseVisualStyleBackColor = true;
			this.button3.Click += new System.EventHandler(this.button3_Click);
			// 
			// button4
			// 
			this.button4.Location = new System.Drawing.Point(287, 59);
			this.button4.Name = "button4";
			this.button4.Size = new System.Drawing.Size(139, 26);
			this.button4.TabIndex = 27;
			this.button4.Text = "Tìm theo mã sinh viên";
			this.button4.UseVisualStyleBackColor = true;
			this.button4.Click += new System.EventHandler(this.button4_Click);
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(794, 127);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(94, 26);
			this.button1.TabIndex = 28;
			this.button1.Text = "Load lại data";
			this.button1.UseVisualStyleBackColor = true;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// khoamotaikhoan
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(925, 385);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button4);
			this.Controls.Add(this.button3);
			this.Controls.Add(this.textBox_timmasinhvien);
			this.Controls.Add(this.textBox_timtendangnhap);
			this.Controls.Add(this.button_motaikhoan);
			this.Controls.Add(this.button_khoataikhoan);
			this.Controls.Add(this.dataGridView_taikhoan);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.Name = "khoamotaikhoan";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Khóa / mở tài khoản";
			((System.ComponentModel.ISupportInitialize)(this.dataGridView_taikhoan)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.DataGridView dataGridView_taikhoan;
		private System.Windows.Forms.Button button_khoataikhoan;
		private System.Windows.Forms.Button button_motaikhoan;
		private System.Windows.Forms.TextBox textBox_timtendangnhap;
		private System.Windows.Forms.TextBox textBox_timmasinhvien;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.Button button4;
		private System.Windows.Forms.Button button1;
	}
}