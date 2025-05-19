namespace ungdungthuviencaocap
{
	partial class chatbot
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
			this.richTextBox_hienthitinnhan = new System.Windows.Forms.RichTextBox();
			this.textBox_nhapcauhoi = new System.Windows.Forms.TextBox();
			this.button_gui = new System.Windows.Forms.Button();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// richTextBox_hienthitinnhan
			// 
			this.richTextBox_hienthitinnhan.BackColor = System.Drawing.SystemColors.ButtonHighlight;
			this.richTextBox_hienthitinnhan.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.richTextBox_hienthitinnhan.Location = new System.Drawing.Point(222, 117);
			this.richTextBox_hienthitinnhan.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.richTextBox_hienthitinnhan.Name = "richTextBox_hienthitinnhan";
			this.richTextBox_hienthitinnhan.ReadOnly = true;
			this.richTextBox_hienthitinnhan.Size = new System.Drawing.Size(872, 426);
			this.richTextBox_hienthitinnhan.TabIndex = 0;
			this.richTextBox_hienthitinnhan.Text = "";
			// 
			// textBox_nhapcauhoi
			// 
			this.textBox_nhapcauhoi.Font = new System.Drawing.Font("Tahoma", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.textBox_nhapcauhoi.Location = new System.Drawing.Point(495, 583);
			this.textBox_nhapcauhoi.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.textBox_nhapcauhoi.Name = "textBox_nhapcauhoi";
			this.textBox_nhapcauhoi.Size = new System.Drawing.Size(571, 27);
			this.textBox_nhapcauhoi.TabIndex = 1;
			this.textBox_nhapcauhoi.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox_nhapcauhoi_KeyDown);
			// 
			// button_gui
			// 
			this.button_gui.Location = new System.Drawing.Point(643, 645);
			this.button_gui.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.button_gui.Name = "button_gui";
			this.button_gui.Size = new System.Drawing.Size(88, 28);
			this.button_gui.TabIndex = 2;
			this.button_gui.Text = "Gửi";
			this.button_gui.UseVisualStyleBackColor = true;
			this.button_gui.Click += new System.EventHandler(this.button_gui_Click);
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(500, 50);
			this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(325, 29);
			this.label1.TabIndex = 3;
			this.label1.Text = "Chatbot giải đáp thắc mắc";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label2.Location = new System.Drawing.Point(272, 583);
			this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(200, 23);
			this.label2.TabIndex = 4;
			this.label2.Text = "Nhập câu hỏi vào đây:";
			// 
			// chatbot
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1194, 823);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.button_gui);
			this.Controls.Add(this.textBox_nhapcauhoi);
			this.Controls.Add(this.richTextBox_hienthitinnhan);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
			this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
			this.Name = "chatbot";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Chatbot QLTV";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.RichTextBox richTextBox_hienthitinnhan;
		private System.Windows.Forms.TextBox textBox_nhapcauhoi;
		private System.Windows.Forms.Button button_gui;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
	}
}