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
using System.IO;
using System.Data.SQLite;
namespace ungdungthuviencaocap
{
	public partial class dangnhap: DevExpress.XtraEditors.XtraForm
	{
        public dangnhap()
		{
            InitializeComponent();
		}
		Modify Modify = new Modify();

		private void Button_dangnhap_Click(object sender, EventArgs e)
		{
			string TenTaiKhoan = textBox_tendangnhap.Text;
			string MatKhau = textBox_matkhau.Text;

			if (TenTaiKhoan.Trim() == "")
			{
				MessageBox.Show("Vui lòng nhập tên tài khoản! ");
				return;
			}

			if (MatKhau.Trim() == "")
			{
				MessageBox.Show("Vui lòng nhập mật khẩu! ");
				return;
			}

			// Sửa câu truy vấn để lấy cả thông tin đăng nhập và trạng thái tài khoản
			string query = "SELECT * FROM TaiKhoan WHERE TenTaiKhoan = @TenTaiKhoan AND MatKhau = @MatKhau";
			SQLiteConnection sqlConnection = null;
			SQLiteCommand command = null;
			SQLiteDataReader reader = null;
			TaiKhoan currentUser = null;
			bool isLocked = false;
			bool loginSuccess = false;

			try
			{
				sqlConnection = Connection.GetSQLiteConnection();
				sqlConnection.Open();
				command = new SQLiteCommand(query, sqlConnection);
				command.Parameters.AddWithValue("@TenTaiKhoan", TenTaiKhoan);
				command.Parameters.AddWithValue("@MatKhau", MatKhau);

				reader = command.ExecuteReader();

				if (reader.Read()) // Nếu tìm thấy người dùng
				{
					// Kiểm tra trạng thái tài khoản
					string trangThai = "active"; // Giá trị mặc định

					// Đọc giá trị trạng thái từ cơ sở dữ liệu
					int trangThaiOrdinal = reader.GetOrdinal("TrangThai");
					if (!reader.IsDBNull(trangThaiOrdinal))
					{
						trangThai = reader.GetString(trangThaiOrdinal);
					}

					// Kiểm tra giá trị trạng thái
					if (trangThai.ToLower() == "lock")
					{
						isLocked = true;
					}
					else
					{
						loginSuccess = true;

						// Tạo đối tượng tài khoản từ dữ liệu đọc được
						currentUser = new TaiKhoan
						{
							MaSinhVien = reader["MaSinhVien"].ToString(),
							TenTaiKhoan = reader["TenTaiKhoan"].ToString(),
							MatKhau = reader["MatKhau"].ToString(),
							Email = reader["Email"].ToString(),
							HoVaTen = reader["HoVaTen"].ToString(),
							Quyen = reader["Quyen"].ToString(),
							DiaChi = reader.IsDBNull(reader.GetOrdinal("DiaChi")) ? "" : reader["DiaChi"].ToString(),
							TrangThai = trangThai
						};

						// Đọc các trường không bắt buộc khác nếu có
						try
						{
							currentUser.Avatar = reader.IsDBNull(reader.GetOrdinal("Avatar")) ? "default.png" : reader["Avatar"].ToString();

							// Kiểm tra từng trường tồn tại trước khi đọc
							if (HasColumn(reader, "SoDienThoai"))
								currentUser.SoDienThoai = reader.IsDBNull(reader.GetOrdinal("SoDienThoai")) ? "" : reader["SoDienThoai"].ToString();

							if (HasColumn(reader, "NgaySinh"))
								currentUser.NgaySinh = reader.IsDBNull(reader.GetOrdinal("NgaySinh")) ? "" : reader["NgaySinh"].ToString();

							if (HasColumn(reader, "GioiTinh"))
								currentUser.GioiTinh = reader.IsDBNull(reader.GetOrdinal("GioiTinh")) ? "" : reader["GioiTinh"].ToString();
						}
						catch (Exception) { /* Bỏ qua lỗi nếu có trường không tồn tại */ }
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi đăng nhập: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			finally
			{
				// Đảm bảo đóng reader, command và connection
				if (reader != null && !reader.IsClosed)
					reader.Close();

				if (command != null)
					command.Dispose();

				if (sqlConnection != null && sqlConnection.State == ConnectionState.Open)
					sqlConnection.Close();
			}

			// Sau khi đã đóng kết nối, xử lý hiển thị thông báo và chuyển form
			if (isLocked)
			{
				MessageBox.Show("Tài khoản của bạn đã bị Khóa, vui lòng liên hệ tới quản lý thư viện để được mở khóa tài khoản.",
					"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			if (loginSuccess && currentUser != null)
			{
				MessageBox.Show("Đăng nhập thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);

				// Lưu thông tin người dùng hiện tại
				UserSession.CurrentUser = currentUser;

				// Kiểm tra vai trò người dùng và chuyển hướng phù hợp
				if (currentUser.Quyen != null && currentUser.Quyen.ToLower() == "admin")
				{
					trangchu adminForm = new trangchu();
					this.Hide();
					adminForm.ShowDialog();
					this.Close();
				}
				else
				{
					nguoidung nguoidung = new nguoidung();
					this.Hide();
					nguoidung.ShowDialog();
					this.Close();
				}
			}
			else if (!isLocked) // Chỉ hiển thị khi không phải tài khoản bị khóa
			{
				MessageBox.Show("Tên tài khoản hoặc mật khẩu không chính xác",
					"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		// Phương thức hỗ trợ để kiểm tra một cột có tồn tại trong kết quả truy vấn hay không
		private bool HasColumn(SQLiteDataReader reader, string columnName)
		{
			for (int i = 0; i < reader.FieldCount; i++)
			{
				if (reader.GetName(i).Equals(columnName, StringComparison.OrdinalIgnoreCase))
					return true;
			}
			return false;
		}




		private void textBox_tendangnhap_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				Button_dangnhap_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}

		private void textBox_matkhau_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				Button_dangnhap_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}

		private void button_dangky_Click(object sender, EventArgs e)
		{
			dangky frmDangKy = new dangky();
			this.Hide();
			frmDangKy.ShowDialog();
			this.Show();
		}

		private void button_quenmatkhau_Click(object sender, EventArgs e)
		{
			quenmatkhau frmquenmatkhau = new quenmatkhau();
			this.Hide();
			frmquenmatkhau.ShowDialog();
			this.Show();
		}

		private void dangnhap_FormClosed(object sender, FormClosedEventArgs e)
		{
			Application.Exit();
		}

		private void dangnhap_FormClosing(object sender, FormClosingEventArgs e)
		{
			if (MessageBox.Show("Bạn có muốn thoát không?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
			{
				e.Cancel = true;
			}
		}
	}
}