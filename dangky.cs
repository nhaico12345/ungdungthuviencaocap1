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
using System.Text.RegularExpressions;

namespace ungdungthuviencaocap
{
	public partial class dangky: DevExpress.XtraEditors.XtraForm
	{
        public dangky()
		{
            InitializeComponent();
		}
		public bool kiemTraTaiKhoan(string ac)
		{
			return Regex.IsMatch(ac, @"^[a-zA-Z0-9_]{6,32}$");
		}
		public bool kiemTraMatKhau(string pw)
		{
			return Regex.IsMatch(pw, @"^[a-zA-Z0-9_]{6,32}$");
		}
		public bool kiemTraEmail(string email)
		{
			return Regex.IsMatch(email, @"^[a-zA-Z0-9_.]{3,20}@gmail.com(.vn|)$");
		}
		Modify Modify = new Modify();
		private void button_dangky_Click(object sender, EventArgs e)
		{
			try
			{
				// Lấy dữ liệu từ form
				string hoTen = textBox_hovaten.Text;
				string tenTaiKhoan = textBox_tendangnhap.Text;
				string email = textBox_email.Text;
				string matKhau = textBox_matkhau.Text;
				string xnMatKhau = textBox_nhaplaimatkhau.Text;
				string maSV = textBox_masv.Text;
				DateTime ngaySinh = dtpBirthday.Value;
				string gioiTinh = combobox_gioitinh.Text;
				string diaChi = textbox_diachi.Text;
				string soDienThoai = textBox_sodienthoai.Text;
				// Kiểm tra các trường bắt buộc
				if (string.IsNullOrWhiteSpace(hoTen))
				{
					MessageBox.Show("Vui lòng nhập họ và tên!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				if (string.IsNullOrWhiteSpace(maSV))
				{
					MessageBox.Show("Vui lòng nhập mã sinh viên!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				if (string.IsNullOrWhiteSpace(gioiTinh))
				{
					MessageBox.Show("Vui lòng chọn giới tính!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				if (string.IsNullOrWhiteSpace(diaChi))
				{
					MessageBox.Show("Vui lòng nhập địa chỉ!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				// Kiểm tra định dạng tên đăng nhập, mật khẩu, email
				if (!kiemTraTaiKhoan(tenTaiKhoan))
				{
					MessageBox.Show("Vui lòng nhập tên đăng nhập dài 6 - 32 kí tự, với các ký tự chữ và số, chữ hoa và chữ thường!");
					return;
				}

				if (!kiemTraMatKhau(matKhau))
				{
					MessageBox.Show("Vui lòng nhập mật khẩu dài 6 - 32 kí tự, với các ký tự chữ và số, chữ hoa và chữ thường!");
					return;
				}

				if (!kiemTraEmail(email))
				{
					MessageBox.Show("Vui lòng nhập email hợp lệ!");
					return;
				}

				if (xnMatKhau != matKhau)
				{
					MessageBox.Show("Mật khẩu không khớp!");
					return;
				}

				// Kiểm tra email đã tồn tại chưa
				if (Modify.taiKhoans("SELECT * FROM TaiKhoan WHERE Email = '" + email + "'").Count != 0)
				{
					MessageBox.Show("Email đã được sử dụng!");
					return;
				}

				// Kiểm tra mã sinh viên đã tồn tại chưa
				if (Modify.taiKhoans("SELECT * FROM TaiKhoan WHERE MaSinhVien = '" + maSV + "'").Count != 0)
				{
					MessageBox.Show("Mã sinh viên đã được đăng ký!");
					return;
				}

				// Định dạng ngày sinh thành chuỗi
				string ngaySinhStr = ngaySinh.ToString("dd/MM/yyyy");

				// Tạo câu lệnh SQL với đầy đủ các trường
				string query = @"INSERT INTO TaiKhoan 
                        (MaSinhVien, TenTaiKhoan, MatKhau, Email, HoVaTen, Avatar, Quyen, DiaChi, SoDienThoai, NgaySinh, GioiTinh, SoDienThoai) 
                        VALUES 
                        (@MaSinhVien, @TenTaiKhoan, @MatKhau, @Email, @HoVaTen, @Avatar, @Quyen, @DiaChi, @SoDienThoai, @NgaySinh, @GioiTinh, @SoDienThoai)";

				// Tạo tham số để ngăn SQL injection
				Dictionary<string, object> parameters = new Dictionary<string, object>
		{
			{ "@MaSinhVien", maSV },
			{ "@TenTaiKhoan", tenTaiKhoan },
			{ "@MatKhau", matKhau },
			{ "@Email", email },
			{ "@HoVaTen", hoTen },
			{ "@Avatar", "default.png" },
			{ "@Quyen", "user" },
			{ "@DiaChi", diaChi },
			{ "@SoDienThoai", soDienThoai },
			{ "@NgaySinh", ngaySinhStr },
			{ "@GioiTinh", gioiTinh }
		};

				// Thực hiện câu lệnh SQL
				int result = Connection.ExecuteNonQuery(query, parameters);

				if (result > 0)
				{
					if (MessageBox.Show("Đăng ký thành công! Bạn có muốn đăng nhập luôn không?",
						"Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information) == DialogResult.Yes)
					{
						this.Close();
					}
				}
				else
				{
					MessageBox.Show("Đăng ký thất bại, vui lòng thử lại sau!",
						"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			catch (Exception ex)
			{
				if (ex.Message.Contains("UNIQUE constraint failed: TaiKhoan.TenTaiKhoan"))
				{
					MessageBox.Show("Tên tài khoản này đã được đăng ký, vui lòng đăng ký tên tài khoản khác!");
				}
				else
				{
					MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		private void textBox_hovaten_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				button_dangky_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}

		private void textBox_tendangnhap_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				button_dangky_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}

		private void textBox_email_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				button_dangky_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}

		private void textBox_matkhau_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				button_dangky_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}

		private void textBox_nhaplaimatkhau_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				button_dangky_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}

		private void textBox_masv_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				button_dangky_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}

		private void textbox_diachi_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				button_dangky_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}

		private void textBox_sodienthoai_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				button_dangky_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}
	}
}