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
	public partial class nguoidung : DevExpress.XtraBars.FluentDesignSystem.FluentDesignForm
	{
		public nguoidung()
		{
			try
			{
				InitializeComponent();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khởi tạo giao diện người dùng: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

		private void mnMuonTraSach_Click(object sender, EventArgs e)
		{
			try
			{
				OpenChildForm(new lichhenmuonsach2());
				lblTieude.Caption = mnMuonTraSach.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form lịch hẹn mượn sách: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnDangXuat_Click(object sender, EventArgs e)
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

		private void mnTrangchu_Click(object sender, EventArgs e)
		{
			try
			{
				if (currentFormChild != null)
				{
					currentFormChild.Close();
				}
				lblTieude.Caption = "Demo thư viện - Tác giả: Trần Đình Đăng";
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi trở về trang chủ: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnHelp_Click(object sender, EventArgs e)
		{
			try
			{
				OpenChildForm(new huongdanquanly());
				lblTieude.Caption = mnHelp.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form hướng dẫn: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void accordionControlElement1_Click(object sender, EventArgs e)
		{
			try
			{
				OpenChildForm(new danhsachsach());
				lblTieude.Caption = mnTonghop.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form danh sách sách: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void mnThongTin_Click(object sender, EventArgs e)
		{
			try
			{
				if (MessageBox.Show("Bảng thông tin sẽ hiện ở một cửa sổ riêng, bạn muốn mở chứ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
				{
					thongtinlienhe thongtinlienhe = new thongtinlienhe();
					thongtinlienhe.ShowDialog();
				}
				lblTieude.Caption = mnThongTin.Text;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form thông tin liên hệ: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void accordionControlElement1_Click_1(object sender, EventArgs e)
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
	}
}
