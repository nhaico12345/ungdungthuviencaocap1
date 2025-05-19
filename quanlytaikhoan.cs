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
using System.Data.SQLite;
using System.Text.RegularExpressions;

namespace ungdungthuviencaocap
{
	public partial class quanlytaikhoan : DevExpress.XtraEditors.XtraForm
	{
		private Modify modify;
		private bool isEditing = false;
		private string currentMaSV = "";

		public quanlytaikhoan()
		{
			InitializeComponent();
			modify = new Modify();
			LoadTaiKhoan(); // Tải dữ liệu tài khoản khi form khởi tạo
			SetControlState(false); // Thiết lập trạng thái ban đầu cho các điều khiển
		}

		// Phương thức thiết lập trạng thái cho các controls
		private void SetControlState(bool enabled)
		{
			textBox_hovaten.Enabled = enabled;
			dtpBirthday.Enabled = enabled;
			textBox_masv.Enabled = enabled ; // Không cho phép sửa mã sinh viên khi đang sửa
			combobox_gioitinh.Enabled = enabled;
			textbox_diachi.Enabled = enabled;
			textBox_tendangnhap.Enabled = enabled ; // Không cho phép sửa tên đăng nhập khi đang sửa
			textBox_email.Enabled = enabled;
			textBox_matkhau.Enabled = enabled;
			textBox_nhaplaimatkhau.Enabled = enabled;
			textBox_sodienthoai.Enabled = enabled;
			comboBox_quyen.Enabled = enabled;

			// Xóa trắng các ô khi không ở chế độ sửa
			if (!enabled && !isEditing)
			{
				ClearForm();
			}
		}

		// Phương thức xóa trắng form
		private void ClearForm()
		{
			textBox_hovaten.Text = "";
			dtpBirthday.Value = DateTime.Now;
			textBox_masv.Text = "";
			combobox_gioitinh.SelectedIndex = -1;
			textbox_diachi.Text = "";
			textBox_tendangnhap.Text = "";
			textBox_email.Text = "";
			textBox_matkhau.Text = "";
			textBox_nhaplaimatkhau.Text = "";
			comboBox_quyen.SelectedIndex = -1;
			textBox_sodienthoai.Text = "";
			currentMaSV = "";
		}

		// Phương thức tải dữ liệu tài khoản
		private void LoadTaiKhoan()
		{
			try
			{
				// Lấy dữ liệu tài khoản từ CSDL
				string query = "SELECT * FROM TaiKhoan admin123";
				DataTable dataTable = Connection.GetData(query);

				// Hiển thị dữ liệu lên DataGridView
				dataGridView_taikhoan.DataSource = dataTable;

				// Điều chỉnh hiển thị các cột
				if (dataGridView_taikhoan.Columns.Contains("MatKhau"))
				{
					dataGridView_taikhoan.Columns["MatKhau"].Visible = false;
				}

				// Đổi tên hiển thị cho các cột
				if (dataGridView_taikhoan.Columns.Contains("MaSinhVien"))
				{
					dataGridView_taikhoan.Columns["MaSinhVien"].HeaderText = "Mã SV";
				}

				if (dataGridView_taikhoan.Columns.Contains("TenTaiKhoan"))
				{
					dataGridView_taikhoan.Columns["TenTaiKhoan"].HeaderText = "Tên đăng nhập";
				}

				if (dataGridView_taikhoan.Columns.Contains("HoVaTen"))
				{
					dataGridView_taikhoan.Columns["HoVaTen"].HeaderText = "Họ và tên";
				}

				if (dataGridView_taikhoan.Columns.Contains("Email"))
				{
					dataGridView_taikhoan.Columns["Email"].HeaderText = "Email";
				}

				if (dataGridView_taikhoan.Columns.Contains("DiaChi"))
				{
					dataGridView_taikhoan.Columns["DiaChi"].HeaderText = "Địa chỉ";
				}

				if (dataGridView_taikhoan.Columns.Contains("NgaySinh"))
				{
					dataGridView_taikhoan.Columns["NgaySinh"].HeaderText = "Ngày sinh";
				}

				if (dataGridView_taikhoan.Columns.Contains("GioiTinh"))
				{
					dataGridView_taikhoan.Columns["GioiTinh"].HeaderText = "Giới tính";
				}

				if (dataGridView_taikhoan.Columns.Contains("Quyen"))
				{
					dataGridView_taikhoan.Columns["Quyen"].HeaderText = "Quyền";
				}

				if (dataGridView_taikhoan.Columns.Contains("SoDienThoai"))
				{
					dataGridView_taikhoan.Columns["SoDienThoai"].HeaderText = "Số điện thoại";
				}

				if (dataGridView_taikhoan.Columns.Contains("Avatar"))
				{
					dataGridView_taikhoan.Columns["Avatar"].Visible = false;
				}

				if (dataGridView_taikhoan.Columns.Contains("TrangThai"))
				{
					dataGridView_taikhoan.Columns["TrangThai"].HeaderText = "Trạng thái";
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi tải dữ liệu tài khoản: " + ex.Message,
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		// Phương thức thêm tài khoản
		private void Button_themtaikhoan_Click(object sender, EventArgs e)
		{
			if (!isEditing)
			{
				// Chuyển sang chế độ thêm mới
				isEditing = true; // Sửa thành true để đánh dấu đang ở chế độ chỉnh sửa
				SetControlState(true);
				ClearForm();
				Button_themtaikhoan.Text = "     Lưu";
			}
			else
			{
				// Đang ở chế độ thêm mới, thực hiện lưu
				try
				{
					// Kiểm tra dữ liệu đầu vào
					if (!ValidateInput())
					{
						return;
					}

					// Kiểm tra mã sinh viên và tên đăng nhập đã tồn tại chưa
					if (CheckMaSVExists(textBox_masv.Text))
					{
						MessageBox.Show("Mã sinh viên đã tồn tại!", "Lỗi",
							MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}

					if (CheckUsernameExists(textBox_tendangnhap.Text))
					{
						MessageBox.Show("Tên đăng nhập đã tồn tại!", "Lỗi",
							MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}

					// Thêm tài khoản mới - đảm bảo tên bảng khớp với cấu trúc DB
					// Sửa lại câu lệnh SQL trong phương thức Button_themtaikhoan_Click
					string query = @"INSERT INTO taikhoan 
                (MaSinhVien, TenTaiKhoan, MatKhau, Email, HoVaTen, 
                 Avatar, Quyen, DiaChi, SoDienThoai, NgaySinh, GioiTinh) 
                VALUES 
                (@MaSinhVien, @TenTaiKhoan, @MatKhau, @Email, @HoVaTen, 
                 @Avatar, @Quyen, @DiaChi, @SoDienThoai, @NgaySinh, @GioiTinh)";


					// Chuẩn bị dữ liệu
					string ngaySinh = dtpBirthday.Value.ToString("yyyy-MM-dd"); // Định dạng ngày tháng chuẩn SQLite

					Dictionary<string, object> parameters = new Dictionary<string, object>
			{
				{ "@MaSinhVien", textBox_masv.Text.Trim() },
				{ "@TenTaiKhoan", textBox_tendangnhap.Text.Trim() },
				{ "@MatKhau", textBox_matkhau.Text },
				{ "@Email", textBox_email.Text.Trim() },
				{ "@HoVaTen", textBox_hovaten.Text.Trim() },
				{ "@Avatar", "default.png" },
				{ "@Quyen", comboBox_quyen.Text },
				{ "@DiaChi", textbox_diachi.Text.Trim() },
				{ "@SoDienThoai", textBox_sodienthoai.Text.Trim() },
                { "@NgaySinh", ngaySinh },
				{ "@GioiTinh", combobox_gioitinh.Text }
			};

					// Thực hiện thêm dữ liệu
					int result = Connection.ExecuteNonQuery(query, parameters);

					if (result > 0)
					{
						MessageBox.Show("Thêm tài khoản thành công!", "Thông báo",
							MessageBoxButtons.OK, MessageBoxIcon.Information);
						LoadTaiKhoan(); // Cập nhật lại danh sách
						SetControlState(false);
						Button_themtaikhoan.Text = "     Thêm";
						isEditing = false;
					}
					else
					{
						MessageBox.Show("Thêm tài khoản thất bại!", "Lỗi",
							MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("Lỗi: " + ex.Message, "Lỗi",
						MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		// Phương thức xóa tài khoản
		private void button_xoataikhoan_Click(object sender, EventArgs e)
		{
			try
			{
				if (dataGridView_taikhoan.SelectedRows.Count == 0)
				{
					MessageBox.Show("Vui lòng chọn tài khoản cần xóa! ", "Thông báo",
						MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				// Lấy thông tin tài khoản được chọn
				string maSinhVien = dataGridView_taikhoan.SelectedRows[0].Cells["MaSinhVien"].Value.ToString();
				string tenTaiKhoan = dataGridView_taikhoan.SelectedRows[0].Cells["TenTaiKhoan"].Value.ToString();

				// Xác nhận xóa
				DialogResult result = MessageBox.Show($"Bạn có chắc chắn muốn xóa tài khoản '{tenTaiKhoan}' không? ",
					"Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

				if (result == DialogResult.Yes)
				{
					// Thực hiện xóa tài khoản
					string query = "DELETE FROM TaiKhoan WHERE MaSinhVien = @MaSinhVien";
					Dictionary<string, object> parameters = new Dictionary<string, object>
					{
						{ "@MaSinhVien", maSinhVien }
					};

					int rowsAffected = Connection.ExecuteNonQuery(query, parameters);

					if (rowsAffected > 0)
					{
						MessageBox.Show("Xóa tài khoản thành công! ", "Thông báo",
							MessageBoxButtons.OK, MessageBoxIcon.Information);
						LoadTaiKhoan(); // Cập nhật lại danh sách 
						ClearForm();
					}
					else
					{
						MessageBox.Show("Không thể xóa tài khoản! ", "Lỗi",
							MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		// Phương thức sửa thông tin tài khoản
		private void button_suathongtintaikhoan_Click(object sender, EventArgs e)
		{
			if (!isEditing)
			{
				// Kiểm tra đã chọn tài khoản chưa
				if (dataGridView_taikhoan.SelectedRows.Count == 0)
				{
					MessageBox.Show("Vui lòng chọn tài khoản cần sửa!", "Thông báo",
						MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				// Lấy thông tin tài khoản được chọn và điền vào form
				currentMaSV = dataGridView_taikhoan.SelectedRows[0].Cells["MaSinhVien"].Value.ToString();
				LoadAccountInfoToForm(currentMaSV);

				// Chuyển sang chế độ sửa
				isEditing = true;
				SetControlState(true);
				button_suathongtintaikhoan.Text = "     Lưu thay đổi";
			}
			else
			{
				// Đang ở chế độ sửa, thực hiện cập nhật
				try
				{
					// Kiểm tra dữ liệu đầu vào
					if (!ValidateInput(true))
					{
						return;
					}

					// Cập nhật thông tin tài khoản
					string query = @"UPDATE TaiKhoan 
                                     SET HoVaTen = @HoVaTen, 
                                         Email = @Email, 
                                         DiaChi = @DiaChi, 
                                         NgaySinh = @NgaySinh, 
                                         GioiTinh = @GioiTinh, 
                                         Quyen = @Quyen,
                                         SoDienThoai= @SoDienThoai";

					// Nếu có cập nhật mật khẩu
					if (!string.IsNullOrEmpty(textBox_matkhau.Text))
					{
						query += ", MatKhau = @MatKhau";
					}

					query += " WHERE MaSinhVien = @MaSinhVien";

					Dictionary<string, object> parameters = new Dictionary<string, object>
					{
						{ "@HoVaTen", textBox_hovaten.Text },
						{ "@Email", textBox_email.Text },
						{ "@DiaChi", textbox_diachi.Text },
						{ "@NgaySinh", dtpBirthday.Value.ToString("dd/MM/yyyy") },
						{ "@GioiTinh", combobox_gioitinh.Text },
						{ "@Quyen", comboBox_quyen.Text },
						{ "@SoDienThoai", textBox_sodienthoai.Text },
						{ "@MaSinhVien", currentMaSV }
					};

					// Thêm tham số mật khẩu nếu cần
					if (!string.IsNullOrEmpty(textBox_matkhau.Text))
					{
						parameters.Add("@MatKhau", textBox_matkhau.Text);
					}

					int result = Connection.ExecuteNonQuery(query, parameters);

					if (result > 0)
					{
						MessageBox.Show("Cập nhật thông tin tài khoản thành công!", "Thông báo",
							MessageBoxButtons.OK, MessageBoxIcon.Information);
						LoadTaiKhoan(); // Cập nhật lại danh sách
						SetControlState(false);
						button_suathongtintaikhoan.Text = "        Sửa thông tin tài khoản";
						isEditing = false;
						ClearForm();
					}
					else
					{
						MessageBox.Show("Cập nhật thông tin tài khoản thất bại!", "Lỗi",
							MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("Lỗi: " + ex.Message, "Lỗi",
						MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		// Phương thức kiểm tra mã sinh viên đã tồn tại chưa
		private bool CheckMaSVExists(string maSV)
		{
			try
			{
				string query = "SELECT COUNT(*) FROM taikhoan WHERE MaSinhVien = @MaSinhVien";
				Dictionary<string, object> parameters = new Dictionary<string, object>
		{
			{ "@MaSinhVien", maSV.Trim() }
		};

				using (SQLiteConnection conn = Connection.GetSQLiteConnection())
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						foreach (var param in parameters)
						{
							cmd.Parameters.AddWithValue(param.Key, param.Value);
						}

						int count = Convert.ToInt32(cmd.ExecuteScalar());
						return count > 0;
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi kiểm tra mã sinh viên: " + ex.Message, "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
		}


		// Phương thức kiểm tra tên đăng nhập đã tồn tại chưa
		private bool CheckUsernameExists(string username)
		{
			try
			{
				string query = "SELECT COUNT(*) FROM taikhoan WHERE TenTaiKhoan = @TenTaiKhoan";
				Dictionary<string, object> parameters = new Dictionary<string, object>
		{
			{ "@TenTaiKhoan", username.Trim() }
		};

				using (SQLiteConnection conn = Connection.GetSQLiteConnection())
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						foreach (var param in parameters)
						{
							cmd.Parameters.AddWithValue(param.Key, param.Value);
						}

						int count = Convert.ToInt32(cmd.ExecuteScalar());
						return count > 0;
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi kiểm tra tên đăng nhập: " + ex.Message, "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
		}

		// Phương thức kiểm tra dữ liệu đầu vào
		private bool ValidateInput(bool isUpdating = false)
		{
			// Kiểm tra họ và tên
			if (string.IsNullOrWhiteSpace(textBox_hovaten.Text))
			{
				MessageBox.Show("Vui lòng nhập họ và tên!", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			// Kiểm tra mã sinh viên nếu không phải đang cập nhật
			if (!isUpdating && string.IsNullOrWhiteSpace(textBox_masv.Text))
			{
				MessageBox.Show("Vui lòng nhập mã sinh viên!", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			// Kiểm tra giới tính
			if (combobox_gioitinh.SelectedIndex == -1)
			{
				MessageBox.Show("Vui lòng chọn giới tính!", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			// Kiểm tra địa chỉ
			if (string.IsNullOrWhiteSpace(textbox_diachi.Text))
			{
				MessageBox.Show("Vui lòng nhập địa chỉ!", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			// Kiểm tra tên đăng nhập nếu không phải đang cập nhật
			if (!isUpdating && string.IsNullOrWhiteSpace(textBox_tendangnhap.Text))
			{
				MessageBox.Show("Vui lòng nhập tên đăng nhập!", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			// Kiểm tra email
			if (string.IsNullOrWhiteSpace(textBox_email.Text))
			{
				MessageBox.Show("Vui lòng nhập email!", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			// Kiểm tra định dạng email
			if (!Regex.IsMatch(textBox_email.Text, @"^[^@\s]+@[^@\s]+\.[^@\s]+$"))
			{
				MessageBox.Show("Email không hợp lệ!", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			// Kiểm tra mật khẩu nếu không phải đang cập nhật hoặc có nhập mật khẩu mới
			if ((!isUpdating && string.IsNullOrWhiteSpace(textBox_matkhau.Text)) ||
				(isUpdating && !string.IsNullOrWhiteSpace(textBox_matkhau.Text) && textBox_matkhau.Text != textBox_nhaplaimatkhau.Text))
			{
				if (!isUpdating)
				{
					MessageBox.Show("Vui lòng nhập mật khẩu!", "Lỗi",
						MessageBoxButtons.OK, MessageBoxIcon.Error);
					return false;
				}
				else if (!string.IsNullOrWhiteSpace(textBox_matkhau.Text))
				{
					MessageBox.Show("Mật khẩu nhập lại không khớp!", "Lỗi",
						MessageBoxButtons.OK, MessageBoxIcon.Error);
					return false;
				}
			}

			// Kiểm tra nhập lại mật khẩu nếu không phải đang cập nhật
			if (!isUpdating && string.IsNullOrWhiteSpace(textBox_nhaplaimatkhau.Text))
			{
				MessageBox.Show("Vui lòng nhập lại mật khẩu!", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			// Kiểm tra mật khẩu và nhập lại mật khẩu khớp nhau
			if (!isUpdating && textBox_matkhau.Text != textBox_nhaplaimatkhau.Text)
			{
				MessageBox.Show("Mật khẩu nhập lại không khớp!", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			// Kiểm tra quyền
			if (comboBox_quyen.SelectedIndex == -1)
			{
				MessageBox.Show("Vui lòng chọn quyền!", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}

			return true;
		}

		// Phương thức tải thông tin tài khoản vào form
		private void LoadAccountInfoToForm(string maSV)
		{
			try
			{
				string query = "SELECT * FROM TaiKhoan WHERE MaSinhVien = @MaSinhVien";
				Dictionary<string, object> parameters = new Dictionary<string, object>
				{
					{ "@MaSinhVien", maSV }
				};

				using (SQLiteConnection conn = Connection.GetSQLiteConnection())
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						foreach (var param in parameters)
						{
							cmd.Parameters.AddWithValue(param.Key, param.Value);
						}

						using (SQLiteDataReader reader = cmd.ExecuteReader())
						{
							if (reader.Read())
							{
								textBox_hovaten.Text = reader["HoVaTen"].ToString();
								textBox_masv.Text = reader["MaSinhVien"].ToString();
								textBox_tendangnhap.Text = reader["TenTaiKhoan"].ToString();
								textBox_email.Text = reader["Email"].ToString();
								textbox_diachi.Text = reader["DiaChi"].ToString();
								textBox_sodienthoai.Text = reader["SoDienThoai"].ToString();
								// Xử lý ngày sinh
								string ngaySinhStr = reader["NgaySinh"].ToString();
								DateTime ngaySinh;
								if (DateTime.TryParse(ngaySinhStr, out ngaySinh))
								{
									dtpBirthday.Value = ngaySinh;
								}

								// Xử lý giới tính
								string gioiTinh = reader["GioiTinh"].ToString();
								if (combobox_gioitinh.Items.Contains(gioiTinh))
								{
									combobox_gioitinh.SelectedItem = gioiTinh;
								}

								// Xử lý quyền
								string quyen = reader["Quyen"].ToString();
								if (comboBox_quyen.Items.Contains(quyen))
								{
									comboBox_quyen.SelectedItem = quyen;
								}

								// Xóa mật khẩu khi sửa
								textBox_matkhau.Text = "";
								textBox_nhaplaimatkhau.Text = "";
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi tải thông tin tài khoản: " + ex.Message,
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		// Xử lý sự kiện click vào DataGridView để xem thông tin tài khoản
		private void dataGridView_taikhoan_CellClick(object sender, DataGridViewCellEventArgs e)
		{
			// Nếu đang ở chế độ chỉnh sửa, không cho phép chọn dòng khác
			if (isEditing)
			{
				return;
			}

			if (e.RowIndex >= 0)
			{
				string maSV = dataGridView_taikhoan.Rows[e.RowIndex].Cells["MaSinhVien"].Value.ToString();
				LoadAccountInfoToForm(maSV);
			}
		}

		private void button_khoamotaikhoan_Click(object sender, EventArgs e)
		{
			khoamotaikhoan khoamotaikhoan = new khoamotaikhoan();
			khoamotaikhoan.ShowDialog();
		}
	}
}