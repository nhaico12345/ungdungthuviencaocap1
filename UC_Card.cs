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
	public partial class UC_Card : DevExpress.XtraEditors.XtraUserControl
	{
		// Các thuộc tính để lưu trữ liên kết và thông tin
		public string FacebookUrl { get; set; } = "https://www.facebook.com/mrsiiiiro112";
		public string TwitterUrl { get; set; } = "https://x.com/CaLivestream";
		public string InstagramUrl { get; set; } = "https://www.instagram.com/dinhdang.2415/";
		public string WhatsAppUrl { get; set; } = "0962938324";
		public string Email { get; set; } = "oosp0305@gmail.com";
		public string PhoneNumber { get; set; } = "0962938324";

		public UC_Card()
		{
			InitializeComponent();
		}

		private void pictureBox_facebook_Click(object sender, EventArgs e)
		{
			System.Diagnostics.Process.Start(FacebookUrl);
		}

		private void pictureBox_gmail_Click(object sender, EventArgs e)
		{
			MessageBox.Show($"Email: {Email}", "Thông tin Gmail", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void pictureBox_twiter_Click(object sender, EventArgs e)
		{
			System.Diagnostics.Process.Start(TwitterUrl);
		}

		private void pictureBox_phone_Click(object sender, EventArgs e)
		{
			MessageBox.Show($"Số điện thoại: {PhoneNumber}", "Thông tin điện thoại", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void pictureBox_instaram_Click(object sender, EventArgs e)
		{
			System.Diagnostics.Process.Start(InstagramUrl);
		}

		private void pictureBox_whatapp_Click(object sender, EventArgs e)
		{
			MessageBox.Show($"Số điện thoại: {PhoneNumber}", "Thông tin điện thoại", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}

		private void artanButton1_Click(object sender, EventArgs e)
		{
			MessageBox.Show("Xin chào, tôi là Trần Đình Đăng\nHọc lớp K16 - CNTT\nKhoa Kỹ thuật - Công nghệ", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
		}
	}
}