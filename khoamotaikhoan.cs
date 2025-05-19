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

namespace ungdungthuviencaocap
{
	public partial class khoamotaikhoan: DevExpress.XtraEditors.XtraForm
	{
        public khoamotaikhoan()
		{
            InitializeComponent();
			LoadTaiKhoan();
		}

		private void button_khoataikhoan_Click(object sender, EventArgs e)
		{
			try
			{
				// Kiểm tra xem đã chọn hàng nào chưa
				if (dataGridView_taikhoan.SelectedRows.Count <= 0)
				{
					MessageBox.Show("Vui lòng chọn tài khoản cần khóa!",
						"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				// Lấy MaSinhVien (khóa chính) của hàng đã chọn
				string maSinhVien = dataGridView_taikhoan.SelectedRows[0].Cells["MaSinhVien"].Value.ToString();

				// Tạo lệnh cập nhật SQL
				string query = "UPDATE TaiKhoan SET TrangThai = 'lock' WHERE MaSinhVien = @MaSinhVien";

				SQLiteConnection sqlConnection = null;
				SQLiteCommand command = null;
				int rowsAffected = 0;
				bool success = false;

				try
				{
					// Mở kết nối và thực hiện cập nhật
					sqlConnection = Connection.GetSQLiteConnection();
					sqlConnection.Open();

					command = new SQLiteCommand(query, sqlConnection);
					command.Parameters.AddWithValue("@MaSinhVien", maSinhVien);

					rowsAffected = command.ExecuteNonQuery();
					success = (rowsAffected > 0);
				}
				catch (Exception ex)
				{
					throw new Exception("Lỗi khi thực hiện khóa tài khoản: " + ex.Message);
				}
				finally
				{
					// Đảm bảo đóng command và connection
					if (command != null)
						command.Dispose();

					if (sqlConnection != null && sqlConnection.State == ConnectionState.Open)
						sqlConnection.Close();
				}

				// Hiển thị thông báo sau khi đã đóng kết nối
				if (success)
				{
					MessageBox.Show("Khóa tài khoản thành công!", "Thông báo",
						MessageBoxButtons.OK, MessageBoxIcon.Information);

					// Làm mới DataGridView sau khi đã hiển thị thông báo
					LoadTaiKhoan();
				}
				else
				{
					MessageBox.Show("Khóa tài khoản thất bại - không tìm thấy tài khoản!",
						"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}


		private void button_motaikhoan_Click(object sender, EventArgs e)
		{
			try
			{
				// Kiểm tra xem đã chọn hàng nào chưa
				if (dataGridView_taikhoan.SelectedRows.Count <= 0)
				{
					MessageBox.Show("Vui lòng chọn tài khoản cần mở khóa!",
						"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				// Lấy MaSinhVien (khóa chính) của hàng đã chọn
				string maSinhVien = dataGridView_taikhoan.SelectedRows[0].Cells["MaSinhVien"].Value.ToString();

				// Tạo lệnh cập nhật SQL
				string query = "UPDATE TaiKhoan SET TrangThai = 'active' WHERE MaSinhVien = @MaSinhVien";

				SQLiteConnection sqlConnection = null;
				SQLiteCommand command = null;
				int rowsAffected = 0;
				bool success = false;

				try
				{
					// Mở kết nối và thực hiện cập nhật
					sqlConnection = Connection.GetSQLiteConnection();
					sqlConnection.Open();

					command = new SQLiteCommand(query, sqlConnection);
					command.Parameters.AddWithValue("@MaSinhVien", maSinhVien);

					rowsAffected = command.ExecuteNonQuery();
					success = (rowsAffected > 0);
				}
				catch (Exception ex)
				{
					throw new Exception("Lỗi khi thực hiện mở khóa tài khoản: " + ex.Message);
				}
				finally
				{
					// Đảm bảo đóng command và connection
					if (command != null)
						command.Dispose();

					if (sqlConnection != null && sqlConnection.State == ConnectionState.Open)
						sqlConnection.Close();
				}

				// Hiển thị thông báo sau khi đã đóng kết nối
				if (success)
				{
					MessageBox.Show("Mở khóa tài khoản thành công!", "Thông báo",
						MessageBoxButtons.OK, MessageBoxIcon.Information);

					// Làm mới DataGridView sau khi đã hiển thị thông báo
					LoadTaiKhoan();
				}
				else
				{
					MessageBox.Show("Mở khóa tài khoản thất bại - không tìm thấy tài khoản!",
						"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}


		private void LoadTaiKhoan()
		{
			try
			{
				// Lấy dữ liệu tài khoản từ CSDL
				string query = "SELECT * FROM TaiKhoan ORDER BY TenTaiKhoan ASC";
				DataTable dataTable = new DataTable();

				SQLiteConnection sqlConnection = null;
				SQLiteDataAdapter dataAdapter = null;

				try
				{
					sqlConnection = Connection.GetSQLiteConnection();
					sqlConnection.Open();

					dataAdapter = new SQLiteDataAdapter(query, sqlConnection);
					dataAdapter.Fill(dataTable);
				}
				finally
				{
					if (dataAdapter != null)
						dataAdapter.Dispose();

					if (sqlConnection != null && sqlConnection.State == ConnectionState.Open)
						sqlConnection.Close();
				}

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
				MessageBox.Show("Lỗi khi tải dữ liệu tài khoản: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}


		private void button3_Click(object sender, EventArgs e)
		{
			try
			{
				string searchText = textBox_timtendangnhap.Text.Trim();

				if (!string.IsNullOrEmpty(searchText))
				{
					// Tạo câu truy vấn SQL với mệnh đề LIKE để tìm kiếm một phần
					string query = "SELECT * FROM TaiKhoan WHERE TenTaiKhoan LIKE @SearchText";

					DataTable dt = new DataTable();
					using (SQLiteConnection sqliteConnection = Connection.GetSQLiteConnection())
					{
						sqliteConnection.Open();
						SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter(query, sqliteConnection);
						dataAdapter.SelectCommand.Parameters.Add("@SearchText", DbType.String).Value = "%" + searchText + "%";
						dataAdapter.Fill(dt);
						sqliteConnection.Close();
					}

					dataGridView_taikhoan.DataSource = dt;
				}
				else
				{
					// Nếu văn bản tìm kiếm trống, tải tất cả tài khoản
					LoadTaiKhoan();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi tìm kiếm theo tên đăng nhập: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button4_Click(object sender, EventArgs e)
		{
			try
			{
				string searchText = textBox_timmasinhvien.Text.Trim();

				if (!string.IsNullOrEmpty(searchText))
				{
					// Tạo câu truy vấn SQL với mệnh đề LIKE để tìm kiếm một phần
					string query = "SELECT * FROM TaiKhoan WHERE MaSinhVien LIKE @SearchText";

					DataTable dt = new DataTable();
					using (SQLiteConnection sqliteConnection = Connection.GetSQLiteConnection())
					{
						sqliteConnection.Open();
						SQLiteDataAdapter dataAdapter = new SQLiteDataAdapter(query, sqliteConnection);
						dataAdapter.SelectCommand.Parameters.Add("@SearchText", DbType.String).Value = "%" + searchText + "%";
						dataAdapter.Fill(dt);
						sqliteConnection.Close();
					}

					dataGridView_taikhoan.DataSource = dt;
				}
				else
				{
					// Nếu văn bản tìm kiếm trống, tải tất cả tài khoản
					LoadTaiKhoan();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi tìm kiếm theo mã sinh viên: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button1_Click(object sender, EventArgs e)
		{
			LoadTaiKhoan();
		}

		private void textBox_timtendangnhap_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				button3_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}

		private void textBox_timmasinhvien_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				button4_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}
	}
}