namespace ungdungthuviencaocap
{
	partial class quetmamuonsach
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
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.label1 = new System.Windows.Forms.Label();
			this.comboBox_camera = new System.Windows.Forms.ComboBox();
			this.xtraScrollableControl1 = new DevExpress.XtraEditors.XtraScrollableControl();
			this.button1 = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.label_hienketqua = new System.Windows.Forms.Label();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
			this.SuspendLayout();
			// 
			// pictureBox1
			// 
			this.pictureBox1.BackColor = System.Drawing.Color.White;
			this.pictureBox1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.pictureBox1.Location = new System.Drawing.Point(0, 98);
			this.pictureBox1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(1099, 511);
			this.pictureBox1.TabIndex = 0;
			this.pictureBox1.TabStop = false;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(43, 37);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(95, 17);
			this.label1.TabIndex = 1;
			this.label1.Text = "Chọn camera:";
			// 
			// comboBox_camera
			// 
			this.comboBox_camera.FormattingEnabled = true;
			this.comboBox_camera.Location = new System.Drawing.Point(154, 37);
			this.comboBox_camera.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.comboBox_camera.Name = "comboBox_camera";
			this.comboBox_camera.Size = new System.Drawing.Size(177, 21);
			this.comboBox_camera.TabIndex = 2;
			// 
			// xtraScrollableControl1
			// 
			this.xtraScrollableControl1.Location = new System.Drawing.Point(551, 0);
			this.xtraScrollableControl1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.xtraScrollableControl1.Name = "xtraScrollableControl1";
			this.xtraScrollableControl1.Size = new System.Drawing.Size(64, 19);
			this.xtraScrollableControl1.TabIndex = 3;
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(368, 35);
			this.button1.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(75, 20);
			this.button1.TabIndex = 4;
			this.button1.Text = "Bắt đầu";
			this.button1.UseVisualStyleBackColor = true;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label2.Location = new System.Drawing.Point(548, 37);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(92, 17);
			this.label2.TabIndex = 5;
			this.label2.Text = "Kết quả scan:";
			// 
			// label_hienketqua
			// 
			this.label_hienketqua.AutoSize = true;
			this.label_hienketqua.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label_hienketqua.Location = new System.Drawing.Point(679, 37);
			this.label_hienketqua.Name = "label_hienketqua";
			this.label_hienketqua.Size = new System.Drawing.Size(89, 17);
			this.label_hienketqua.TabIndex = 6;
			this.label_hienketqua.Text = "Hiện kết quả:";
			// 
			// quetmamuonsach
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1099, 609);
			this.Controls.Add(this.label_hienketqua);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.xtraScrollableControl1);
			this.Controls.Add(this.comboBox_camera);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.pictureBox1);
			this.IconOptions.Image = global::ungdungthuviencaocap.Properties.Resources.qr;
			this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.MaximizeBox = false;
			this.Name = "quetmamuonsach";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Quét mã mượn sách";
			this.Load += new System.EventHandler(this.quetmamuonsach_Load);
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ComboBox comboBox_camera;
		private DevExpress.XtraEditors.XtraScrollableControl xtraScrollableControl1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label_hienketqua;
	}
}