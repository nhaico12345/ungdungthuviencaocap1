namespace ungdungthuviencaocap
{
	partial class theloai
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
            this.button_lammoi = new System.Windows.Forms.Button();
            this.button_sua = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox_tentheloai = new System.Windows.Forms.TextBox();
            this.dataGridView_theloai = new System.Windows.Forms.DataGridView();
            this.panel3 = new System.Windows.Forms.Panel();
            this.dataGridView_tacpham = new System.Windows.Forms.DataGridView();
            this.label11 = new System.Windows.Forms.Label();
            this.textBox_trangthai = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.textBox_mota = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox_matheloai = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_theloai)).BeginInit();
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
            this.label10.Size = new System.Drawing.Size(139, 19);
            this.label10.TabIndex = 0;
            this.label10.Text = "Danh sách thể loại";
            // 
            // button_loc
            // 
            this.button_loc.Location = new System.Drawing.Point(217, 271);
            this.button_loc.Name = "button_loc";
            this.button_loc.Size = new System.Drawing.Size(65, 23);
            this.button_loc.TabIndex = 26;
            this.button_loc.Text = "Lọc";
            this.button_loc.UseVisualStyleBackColor = true;
            // 
            // button_lammoi
            // 
            this.button_lammoi.Location = new System.Drawing.Point(397, 271);
            this.button_lammoi.Name = "button_lammoi";
            this.button_lammoi.Size = new System.Drawing.Size(72, 23);
            this.button_lammoi.TabIndex = 22;
            this.button_lammoi.Text = "Làm mới";
            this.button_lammoi.UseVisualStyleBackColor = true;
            this.button_lammoi.Click += new System.EventHandler(this.button_lammoi_Click);
            // 
            // button_sua
            // 
            this.button_sua.Location = new System.Drawing.Point(309, 271);
            this.button_sua.Name = "button_sua";
            this.button_sua.Size = new System.Drawing.Size(71, 23);
            this.button_sua.TabIndex = 20;
            this.button_sua.Text = "Cập nhật";
            this.button_sua.UseVisualStyleBackColor = true;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(296, 112);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(90, 18);
            this.label2.TabIndex = 1;
            this.label2.Text = "Tên thể loại:";
            // 
            // textBox_tentheloai
            // 
            this.textBox_tentheloai.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.textBox_tentheloai.Location = new System.Drawing.Point(217, 142);
            this.textBox_tentheloai.Name = "textBox_tentheloai";
            this.textBox_tentheloai.Size = new System.Drawing.Size(252, 26);
            this.textBox_tentheloai.TabIndex = 4;
            // 
            // dataGridView_theloai
            // 
            this.dataGridView_theloai.AllowUserToAddRows = false;
            this.dataGridView_theloai.AllowUserToDeleteRows = false;
            this.dataGridView_theloai.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridView_theloai.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.dataGridView_theloai.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView_theloai.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dataGridView_theloai.Location = new System.Drawing.Point(0, 31);
            this.dataGridView_theloai.Name = "dataGridView_theloai";
            this.dataGridView_theloai.ReadOnly = true;
            this.dataGridView_theloai.Size = new System.Drawing.Size(674, 290);
            this.dataGridView_theloai.TabIndex = 1;
            this.dataGridView_theloai.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_theloai_CellClick);
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
            this.label11.Location = new System.Drawing.Point(548, 9);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(152, 19);
            this.label11.TabIndex = 0;
            this.label11.Text = "Danh sách tác phẩm";
            // 
            // textBox_trangthai
            // 
            this.textBox_trangthai.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.textBox_trangthai.Location = new System.Drawing.Point(217, 211);
            this.textBox_trangthai.Name = "textBox_trangthai";
            this.textBox_trangthai.Size = new System.Drawing.Size(252, 26);
            this.textBox_trangthai.TabIndex = 16;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(306, 181);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(80, 18);
            this.label9.TabIndex = 15;
            this.label9.Text = "Trạng thái:";
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.panel2.Controls.Add(this.dataGridView_theloai);
            this.panel2.Controls.Add(this.label10);
            this.panel2.Location = new System.Drawing.Point(-1, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(674, 321);
            this.panel2.TabIndex = 4;
            // 
            // textBox_mota
            // 
            this.textBox_mota.Location = new System.Drawing.Point(3, 75);
            this.textBox_mota.Multiline = true;
            this.textBox_mota.Name = "textBox_mota";
            this.textBox_mota.Size = new System.Drawing.Size(195, 162);
            this.textBox_mota.TabIndex = 17;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(187, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(135, 19);
            this.label1.TabIndex = 0;
            this.label1.Text = "Thông tin thể loại";
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)));
            this.panel1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.textBox_matheloai);
            this.panel1.Controls.Add(this.button_loc);
            this.panel1.Controls.Add(this.button_lammoi);
            this.panel1.Controls.Add(this.button_sua);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.textBox_tentheloai);
            this.panel1.Controls.Add(this.textBox_trangthai);
            this.panel1.Controls.Add(this.textBox_mota);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(676, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(483, 321);
            this.panel1.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(302, 44);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(84, 18);
            this.label3.TabIndex = 27;
            this.label3.Text = "Mã thể loại:";
            // 
            // textBox_matheloai
            // 
            this.textBox_matheloai.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(163)));
            this.textBox_matheloai.Location = new System.Drawing.Point(217, 75);
            this.textBox_matheloai.Name = "textBox_matheloai";
            this.textBox_matheloai.Size = new System.Drawing.Size(252, 26);
            this.textBox_matheloai.TabIndex = 28;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(79, 44);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(46, 18);
            this.label4.TabIndex = 29;
            this.label4.Text = "Mô tả";
            // 
            // theloai
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1159, 677);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.MaximizeBox = false;
            this.Name = "theloai";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Thể loại";
            this.Load += new System.EventHandler(this.theloai_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView_theloai)).EndInit();
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
		private System.Windows.Forms.Button button_lammoi;
		private System.Windows.Forms.Button button_sua;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.TextBox textBox_tentheloai;
		private System.Windows.Forms.DataGridView dataGridView_theloai;
		private System.Windows.Forms.Panel panel3;
		private System.Windows.Forms.DataGridView dataGridView_tacpham;
		private System.Windows.Forms.Label label11;
		private System.Windows.Forms.TextBox textBox_trangthai;
		private System.Windows.Forms.Label label9;
		private System.Windows.Forms.Panel panel2;
		private System.Windows.Forms.TextBox textBox_mota;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox textBox_matheloai;
		private System.Windows.Forms.Label label4;
	}
}