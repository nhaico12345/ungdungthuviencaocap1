using DevExpress.XtraBars;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ungdungthuviencaocap
{
	public partial class trangchu : DevExpress.XtraBars.FluentDesignSystem.FluentDesignForm
	{
		public trangchu()
		{
			try
			{
				InitializeComponent();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khởi tạo form trang chủ: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private Form currentFormChild;
		private void OpenChildForm(Form childForm)
		{
			try
			{
				if (currentFormChild != null)
				{
					currentFormChild.Close();
				}
				currentFormChild = childForm;
				childForm.TopLevel = false;
				childForm.FormBorderStyle = FormBorderStyle.None;
				childForm.Dock = DockStyle.Fill;
				fluentDesignFormContainer1.Controls.Add(childForm);
				fluentDesignFormContainer1.Tag = childForm;
				childForm.BringToFront();
				childForm.Show();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form con: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnQuanlysach_Click(object sender, EventArgs e)
		{
			try
			{
				OpenChildForm(new quanlysach());
				lblTieude.Caption = mnQuanlysach.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form quản lý sách: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnMuontrasach_Click(object sender, EventArgs e)
		{
			try
			{
				OpenChildForm(new muontrasach());
				lblTieude.Caption = mnMuontrasach.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form mượn trả sách: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnTrangchu_Click(object sender, EventArgs e)
		{
			try
			{
				if (currentFormChild != null)
				{
					currentFormChild.Close();
				}
				lblTieude.Caption = "Demo quản lý thư viện - Tác giả: Trần Đình Đăng";
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi trở về trang chủ: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnThongtin_Click(object sender, EventArgs e)
		{
			try
			{
				OpenChildForm(new thongtinlienhe());
				lblTieude.Caption = mnThongtin.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form thông tin liên hệ: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnDangxuat_Click(object sender, EventArgs e)
		{
			try
			{
				if (MessageBox.Show("Bạn có muốn đăng xuất không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
				{
					this.Hide();
					dangnhap dangnhap = new dangnhap();
					dangnhap.ShowDialog();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi đăng xuất: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnBangthongke_Click(object sender, EventArgs e)
		{
			try
			{
				if (MessageBox.Show("Bảng thống kê sẽ hiện ở một cửa sổ riêng, bạn muốn mở chứ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
				{
					bangthongke bangthongke = new bangthongke();
					bangthongke.ShowDialog();
				}
				lblTieude.Caption = mnBangthongke.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở bảng thống kê: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnHuongdan_Click(object sender, EventArgs e)
		{
			try
			{
				OpenChildForm(new huongdanquanly());
				lblTieude.Caption = mnHuongdan.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form hướng dẫn quản lý: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnThongTinTaiKhoan_Click(object sender, EventArgs e)
		{
			try
			{
				OpenChildForm(new thongtintaikhoan());
				lblTieude.Caption = mnThongTinTaiKhoan.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form thông tin tài khoản: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnQuanlytaikhoan_Click(object sender, EventArgs e)
		{
			try
			{
				OpenChildForm(new quanlytaikhoan());
				lblTieude.Caption = mnQuanlytaikhoan.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form quản lý tài khoản: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnCaidat_Click(object sender, EventArgs e)
		{
			try
			{
				OpenChildForm(new caidat());
				lblTieude.Caption = mnCaidat.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form quản lý tài khoản: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void accordionControl1_Click(object sender, EventArgs e)
		{

		}

		public void ChangeAccordionControlColor(System.Drawing.Color color)
		{
			if (accordionControl1 != null)
			{
				accordionControl1.Appearance.AccordionControl.BackColor = color;
				// Bạn cũng có thể muốn cập nhật các thuộc tính giao diện khác để đảm bảo tính nhất quán
				accordionControl1.Refresh();
			}
		}

		private void accordionControlElement1_Click(object sender, EventArgs e)
		{
			try
			{
				OpenChildForm(new chatbot());
				lblTieude.Caption = mnChatbot.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form quản lý tài khoản: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnThongtinsach_Click(object sender, EventArgs e)
		{
			try
			{
				OpenChildForm(new timkiemsach());
				lblTieude.Caption = mnThongtinsach.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form quản lý tài khoản: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnTacgia_Click(object sender, EventArgs e)
		{
			try
			{
				OpenChildForm(new tacgia());
				lblTieude.Caption = mnThongtinsach.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form quản lý tài khoản: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnTheloai_Click(object sender, EventArgs e)
		{
			try
			{
				OpenChildForm(new theloai());
				lblTieude.Caption = mnTheloai.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form quản lý tài khoản: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnNhaxuatban_Click(object sender, EventArgs e)
		{
			try
			{
				OpenChildForm(new nhaxuatban());
				lblTieude.Caption = mnNhaxuatban.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form quản lý tài khoản: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
