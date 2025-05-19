using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace ungdungthuviencaocap
{
	public partial class UC_Card3: DevExpress.XtraEditors.XtraUserControl
	{
		// Các thuộc tính để lưu trữ liên kết và thông tin
		public string FacebookUrl { get; set; } = "https://www.facebook.com/thai.bao.578542";
		public string TwitterUrl { get; set; } = "Chưa có thông tin";
		public string InstagramUrl { get; set; } = "Chưa có thông tin";
		public string WhatsAppUrl { get; set; } = "0912798369";
		public string Email { get; set; } = "Chưa có thông tin";
		public string PhoneNumber { get; set; } = "0912798369";
		public UC_Card3()
		{
            InitializeComponent();
		}

		private void pictureBox_facebook_Click(object sender, EventArgs e)
		{
			System.Diagnostics.Process.Start(FacebookUrl);
		}

		private void pictureBox_gmail_Click(object sender, EventArgs e)
		{
			MessageBox.Show($"Gmail: {Email}", "Thông tin Gmail", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void pictureBox_twiter_Click(object sender, EventArgs e)
		{
			MessageBox.Show($"X: {TwitterUrl}", "Thông tin X", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void pictureBox_phone_Click(object sender, EventArgs e)
		{
			MessageBox.Show($"Số điện thoại: {PhoneNumber}", "Thông tin điện thoại", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void pictureBox_instaram_Click(object sender, EventArgs e)
		{
			MessageBox.Show($"Instagram: {InstagramUrl}", "Thông tin Instagram", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void pictureBox_whatapp_Click(object sender, EventArgs e)
		{
			MessageBox.Show($"WhatsApp: {WhatsAppUrl}", "Thông tin WhatsApp", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void artanButton1_Click(object sender, EventArgs e)
		{
			MessageBox.Show("Xin chào, tôi là Thái Lê Gia Bảo\nHọc lớp K16 - CNTT\nKhoa Kỹ thuật - Công nghệ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
	}
}
