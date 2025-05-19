using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace ungdungthuviencaocap
{
	public partial class quenmatkhau : DevExpress.XtraEditors.XtraForm
	{
		public quenmatkhau()
		{
			InitializeComponent();
			label2.Text = "";
			textBox_nhapmk.Visible = false;
			textBox_nhaplaimk.Visible = false;
			label4.Visible = false;
			label5.Visible = false;
			button_doimk.Visible = false; // Ẩn nút đổi mật khẩu ban đầu
		}

		Modify Modify = new Modify();

		private void button_laylaimatkhau_Click(object sender, EventArgs e)
		{
			try
			{
				string email = textBox_email.Text.Trim();
				string tenDangNhap = textBox_tendangnhap.Text.Trim();

				if (string.IsNullOrWhiteSpace(email) || string.IsNullOrWhiteSpace(tenDangNhap))
				{
					MessageBox.Show("Vui lòng nhập tên đăng nhập và email đã đăng kí!",
						"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				string query = "SELECT * FROM TaiKhoan WHERE Email = @Email AND TenTaiKhoan = @TenTaiKhoan";
				Dictionary<string, object> parameters = new Dictionary<string, object>
				{
					{ "@Email", email },
					{ "@TenTaiKhoan", tenDangNhap }
				};

				// Sửa lại cách lấy dữ liệu để sử dụng parameters
				var taiKhoans = Modify.taiKhoans(query, parameters);

				if (taiKhoans != null && taiKhoans.Count > 0)
				{
					label2.ForeColor = Color.Green;
					label2.Text = "Mật khẩu của bạn là: " + taiKhoans[0].MatKhau;

					var result = MessageBox.Show("Mật khẩu của bạn đã được hiển thị. Bạn có muốn đổi mật khẩu không?",
						"Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

					if (result == DialogResult.Yes)
					{
						textBox_nhapmk.Visible = true;
						textBox_nhaplaimk.Visible = true;
						label4.Visible = true;
						label5.Visible = true;
						button_doimk.Visible = true;
					}
				}
				else
				{
					label2.ForeColor = Color.Red;
					label2.Text = "Tên đăng nhập hoặc email không tồn tại!";
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button_doimk_Click(object sender, EventArgs e)
		{
			try
			{
				string newPassword = textBox_nhapmk.Text;
				string confirmPassword = textBox_nhaplaimk.Text;

				if (string.IsNullOrWhiteSpace(newPassword) || string.IsNullOrWhiteSpace(confirmPassword))
				{
					MessageBox.Show("Hãy nhập đầy đủ thông tin!");
					return;
				}

				if (newPassword != confirmPassword)
				{
					MessageBox.Show("Mật khẩu không giống nhau, hãy kiểm tra lại!");
					return;
				}

				// Cập nhật mật khẩu sử dụng parameters
				string query = "UPDATE TaiKhoan SET MatKhau = @MatKhau WHERE TenTaiKhoan = @TenTaiKhoan AND Email = @Email";
				Dictionary<string, object> parameters = new Dictionary<string, object>
				{
					{ "@MatKhau", newPassword },
					{ "@TenTaiKhoan", textBox_tendangnhap.Text.Trim() },
					{ "@Email", textBox_email.Text.Trim() }
				};

				int result = Connection.ExecuteNonQuery(query, parameters);

				if (result > 0)
				{
					MessageBox.Show("Thay đổi mật khẩu thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);

					// Reset form về trạng thái ban đầu
					textBox_nhapmk.Visible = false;
					textBox_nhaplaimk.Visible = false;
					label4.Visible = false;
					label5.Visible = false;
					button_doimk.Visible = false;

					// Cập nhật lại thông báo mật khẩu
					string queryGetNew = "SELECT * FROM TaiKhoan WHERE TenTaiKhoan = @TenTaiKhoan AND Email = @Email";
					var updatedUser = Modify.taiKhoans(queryGetNew, parameters);

					if (updatedUser != null && updatedUser.Count > 0)
					{
						label2.Text = "Mật khẩu mới của bạn là: " + updatedUser[0].MatKhau;
					}
				}
				else
				{
					MessageBox.Show("Thay đổi mật khẩu thất bại!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void textBox_tendangnhap_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				button_laylaimatkhau_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}

		private void textBox_email_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				button_laylaimatkhau_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}
	}
}