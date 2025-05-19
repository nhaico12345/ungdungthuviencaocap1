namespace ungdungthuviencaocap
{
	partial class caidat
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
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(caidat));
			this.label1 = new System.Windows.Forms.Label();
			this.colorPickEdit1 = new DevExpress.XtraEditors.ColorPickEdit();
			this.button_thaymau = new System.Windows.Forms.Button();
			this.button_dongbolen = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.button_dongboxuong = new System.Windows.Forms.Button();
			this.button_kiemtraketnoi = new System.Windows.Forms.Button();
			((System.ComponentModel.ISupportInitialize)(this.colorPickEdit1.Properties)).BeginInit();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(54, 52);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(162, 19);
			this.label1.TabIndex = 0;
			this.label1.Text = "Chỉnh màu giao diện:";
			// 
			// colorPickEdit1
			// 
			this.colorPickEdit1.EditValue = System.Drawing.Color.Empty;
			this.colorPickEdit1.Location = new System.Drawing.Point(239, 54);
			this.colorPickEdit1.Name = "colorPickEdit1";
			this.colorPickEdit1.Properties.AutomaticColor = System.Drawing.Color.Black;
			this.colorPickEdit1.Properties.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
			this.colorPickEdit1.Size = new System.Drawing.Size(100, 20);
			this.colorPickEdit1.TabIndex = 1;
			// 
			// button_thaymau
			// 
			this.button_thaymau.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button_thaymau.Location = new System.Drawing.Point(356, 50);
			this.button_thaymau.Name = "button_thaymau";
			this.button_thaymau.Size = new System.Drawing.Size(72, 28);
			this.button_thaymau.TabIndex = 2;
			this.button_thaymau.Text = "Thay màu";
			this.button_thaymau.UseVisualStyleBackColor = true;
			this.button_thaymau.Click += new System.EventHandler(this.button_thaymau_Click);
			// 
			// button_dongbolen
			// 
			this.button_dongbolen.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button_dongbolen.Location = new System.Drawing.Point(662, 56);
			this.button_dongbolen.Name = "button_dongbolen";
			this.button_dongbolen.Size = new System.Drawing.Size(190, 28);
			this.button_dongbolen.TabIndex = 4;
			this.button_dongbolen.Text = "Đồng bộ dữ liệu lên server";
			this.button_dongbolen.UseVisualStyleBackColor = true;
			this.button_dongbolen.Click += new System.EventHandler(this.button_dongbolen_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label2.Location = new System.Drawing.Point(747, 24);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(224, 19);
			this.label2.TabIndex = 5;
			this.label2.Text = "Đồng bộ file dữ liệu với server";
			// 
			// button_dongboxuong
			// 
			this.button_dongboxuong.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button_dongboxuong.Location = new System.Drawing.Point(871, 56);
			this.button_dongboxuong.Name = "button_dongboxuong";
			this.button_dongboxuong.Size = new System.Drawing.Size(211, 28);
			this.button_dongboxuong.TabIndex = 6;
			this.button_dongboxuong.Text = "Đồng bộ dữ liệu từ server xuống";
			this.button_dongboxuong.UseVisualStyleBackColor = true;
			this.button_dongboxuong.Click += new System.EventHandler(this.button_dongboxuong_Click);
			// 
			// button_kiemtraketnoi
			// 
			this.button_kiemtraketnoi.BackColor = System.Drawing.Color.White;
			this.button_kiemtraketnoi.Cursor = System.Windows.Forms.Cursors.Hand;
			this.button_kiemtraketnoi.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.button_kiemtraketnoi.Image = global::ungdungthuviencaocap.Properties.Resources.link___Copy;
			this.button_kiemtraketnoi.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
			this.button_kiemtraketnoi.Location = new System.Drawing.Point(58, 110);
			this.button_kiemtraketnoi.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.button_kiemtraketnoi.Name = "button_kiemtraketnoi";
			this.button_kiemtraketnoi.Size = new System.Drawing.Size(177, 42);
			this.button_kiemtraketnoi.TabIndex = 21;
			this.button_kiemtraketnoi.Text = "        Kiểm tra kết nối";
			this.button_kiemtraketnoi.UseVisualStyleBackColor = false;
			this.button_kiemtraketnoi.Click += new System.EventHandler(this.button_kiemtraketnoi_Click);
			// 
			// caidat
			// 
			this.Appearance.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(255)))), ((int)(((byte)(255)))));
			this.Appearance.Options.UseBackColor = true;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1140, 566);
			this.Controls.Add(this.button_kiemtraketnoi);
			this.Controls.Add(this.button_dongboxuong);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.button_dongbolen);
			this.Controls.Add(this.button_thaymau);
			this.Controls.Add(this.colorPickEdit1);
			this.Controls.Add(this.label1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			this.IconOptions.LargeImage = ((System.Drawing.Image)(resources.GetObject("caidat.IconOptions.LargeImage")));
			this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
			this.Name = "caidat";
			this.Text = "Cài đặt";
			((System.ComponentModel.ISupportInitialize)(this.colorPickEdit1.Properties)).EndInit();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label label1;
		private DevExpress.XtraEditors.ColorPickEdit colorPickEdit1;
		private System.Windows.Forms.Button button_thaymau;
		private System.Windows.Forms.Button button_dongbolen;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button button_dongboxuong;
		private System.Windows.Forms.Button button_kiemtraketnoi;
	}
}