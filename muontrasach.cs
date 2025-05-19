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
using System.Data.SqlClient;
using System.Data.SQLite;
using DevExpress.XtraEditors.Filtering.Templates;

namespace ungdungthuviencaocap
{
	public partial class muontrasach : DevExpress.XtraEditors.XtraForm
	{
		public muontrasach()
		{
			InitializeComponent();
		}
		Modify modify;
		private void LoadDataGridView()
		{
			try
			{
				// Chỉ hiển thị các phiếu có IsActive = 1 (chưa trả)
				string query = "SELECT * FROM muontrasach WHERE IsActive = 1";

				using (SQLiteConnection con = Connection.GetSQLiteConnection())
				{
					con.Open();
					SQLiteCommand command = new SQLiteCommand(query, con);
					SQLiteDataAdapter adapter = new SQLiteDataAdapter(command);
					DataTable dt = new DataTable();
					adapter.Fill(dt);
					dataGridView1.DataSource = dt;
					dataGridView1.Columns["ID"].Visible = false;
					dataGridView1.Columns["IsActive"].Visible = false;
					dataGridView1.Columns["TenSach"].HeaderText = "Tên sách";
					dataGridView1.Columns["MaSinhVien"].HeaderText = "Mã sinh viên";
					dataGridView1.Columns["HoVaTen"].HeaderText = "Họ và tên";
					dataGridView1.Columns["LoaiPhieu"].HeaderText = "Loại phiếu";
					dataGridView1.Columns["NgayMuon"].HeaderText = "Ngày mượn";
					dataGridView1.Columns["NgayTra"].HeaderText = "Ngày trả";
					dataGridView1.Columns["SoLuong"].HeaderText = "Số lượng";
					// Đảm bảo hiển thị MaSach
					if (dataGridView1.Columns["MaSach"] != null)
					{
						dataGridView1.Columns["MaSach"].HeaderText = "Mã Sách";
						dataGridView1.Columns["MaSach"].DisplayIndex = 1; // Điều chỉnh vị trí hiển thị nếu cần
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}


		private void muontrasach_Load(object sender, EventArgs e)
		{
			modify = new Modify();
			LoadDataGridView();
			DataTable dt = new DataTable();
			using (SQLiteConnection con= Connection.GetSQLiteConnection())
			{
				con.Open();
				SQLiteCommand command = new SQLiteCommand();
				command.CommandText = "select * from quanlysach ORDER BY TenSach ASC";
				command.Connection = con;
				SQLiteDataReader readSQL = command.ExecuteReader();
				dt.Load(readSQL);
				comboBox_tensach.DataSource = dt;
				comboBox_tensach.DisplayMember = "Tensach";
				comboBox_tensach.ValueMember = "tensach";
			}
		}
		private void button_phieumuon_Click(object sender, EventArgs e)
		{
			try
			{
				string Tensach = this.comboBox_tensach.Text;
				string masinhvien = this.textBox_masinhvien.Text;
				string Hovaten = this.textBox_hovaten.Text;
				string Loaiphieu = this.comboBox_LoaiPhieu.Text;
				DateTime ngaymuon = dateTimePicker_muonsach.Value;
				DateTime ngaytra = dateTimePicker_trasach.Value;
				string soluong = this.textBox_soluong.Text;

				// Lấy mã sách từ textBox và kiểm tra
				string masach = this.textBox_masach.Text.Trim();
				if (string.IsNullOrEmpty(masach))
				{
					MessageBox.Show("Mã sách không được để trống! ", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				// Kiểm tra đầu vào số lượng
				int soLuongInt;
				if (!int.TryParse(soluong, out soLuongInt) || soLuongInt <= 0)
				{
					MessageBox.Show("Số lượng sách mượn phải là số nguyên dương! ", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				// Kiểm tra xem sách có tồn tại và đủ số lượng
				using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
				{
					sqlConnection.Open();
					string checkQuery = "SELECT SoLuong FROM quanlysach WHERE TenSach = @TenSach";
					SQLiteCommand checkCommand = new SQLiteCommand(checkQuery, sqlConnection);
					checkCommand.Parameters.AddWithValue("@TenSach", Tensach);

					object result = checkCommand.ExecuteScalar();
					if (result == null)
					{
						MessageBox.Show("Không tìm thấy sách này trong hệ thống! ", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}

					int availableQuantity = Convert.ToInt32(result);

					if (soLuongInt > availableQuantity)
					{
						MessageBox.Show("Sách không đủ, hãy đợi thư viện bổ sung hoặc người khác trả sách nhé! ", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}

					// Cập nhật số lượng sách
					string updateQuery = "UPDATE quanlysach SET SoLuong = SoLuong - @SoLuongMuon WHERE TenSach = @TenSach";
					SQLiteCommand updateCommand = new SQLiteCommand(updateQuery, sqlConnection);
					updateCommand.Parameters.AddWithValue("@SoLuongMuon", soLuongInt);
					updateCommand.Parameters.AddWithValue("@TenSach", Tensach);
					updateCommand.ExecuteNonQuery();
				}

				// Tạo phiếu mượn có thêm tham số mã sách
				muontra muontra = new muontra(0, Tensach, masinhvien, Hovaten, Loaiphieu, ngaymuon, ngaytra, soluong, masach);

				if (modify.themphieu(muontra))
				{
					// Cập nhật thống kê mượn
					modify.CapNhatThongKeMuon(masinhvien, Hovaten, soLuongInt, ngaymuon);
					MessageBox.Show("Thêm phiếu mượn thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					LoadDataGridView();
				}
				else
				{
					// Nếu thêm phiếu thất bại, khôi phục số lượng sách
					using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
					{
						sqlConnection.Open();
						string restoreQuery = "UPDATE quanlysach SET SoLuong = SoLuong + @SoLuongMuon WHERE TenSach = @TenSach";
						SQLiteCommand restoreCommand = new SQLiteCommand(restoreQuery, sqlConnection);
						restoreCommand.Parameters.AddWithValue("@SoLuongMuon", soLuongInt);
						restoreCommand.Parameters.AddWithValue("@TenSach", Tensach);
						restoreCommand.ExecuteNonQuery();
					}

					MessageBox.Show("Thêm phiếu mượn thất bại", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}




		private void button_sua_Click(object sender, EventArgs e)
		{
			try
			{
				int ID = Convert.ToInt32(dataGridView1.CurrentRow.Cells["ID"].Value);
				string Tensach = this.comboBox_tensach.Text;
				string masinhvien = this.textBox_masinhvien.Text;
				string Hovaten = this.textBox_hovaten.Text;
				string Loaiphieu = this.comboBox_LoaiPhieu.Text;
				DateTime ngaymuon = dateTimePicker_muonsach.Value;
				DateTime ngaytra = dateTimePicker_trasach.Value;
				string soluong = this.textBox_soluong.Text;
				string masach = this.textBox_masach.Text;

				muontra muontra = new muontra(ID, Tensach, masinhvien, Hovaten, Loaiphieu, ngaymuon, ngaytra, soluong, masach);
				if (modify.suaphieu(muontra))
				{
					MessageBox.Show("Sửa phiếu mượn thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					LoadDataGridView();
				}
				else
				{
					MessageBox.Show("Sửa phiếu mượn thất bại", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}


		private void button_trasach_Click(object sender, EventArgs e)
		{
			try
			{
				if (dataGridView1.SelectedRows.Count == 0)
				{
					MessageBox.Show("Bạn chưa chọn sách để trả", "Lỗi",
						MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				int ID = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells["ID"].Value);
				string masinhvien = dataGridView1.SelectedRows[0].Cells["MaSinhVien"].Value.ToString();
				string hoVaTen = dataGridView1.SelectedRows[0].Cells["HoVaTen"].Value.ToString();
				string tenSach = dataGridView1.SelectedRows[0].Cells["TenSach"].Value.ToString();
				int soLuong = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells["SoLuong"].Value);
				DateTime ngayTra = DateTime.Now;

				// Cập nhật số lượng sách
				using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
				{
					sqlConnection.Open();
					string updateBookQuery = "UPDATE quanlysach SET SoLuong = SoLuong + @SoLuongTra WHERE TenSach = @TenSach";
					SQLiteCommand updateBookCommand = new SQLiteCommand(updateBookQuery, sqlConnection);
					updateBookCommand.Parameters.AddWithValue("@SoLuongTra", soLuong);
					updateBookCommand.Parameters.AddWithValue("@TenSach", tenSach);
					updateBookCommand.ExecuteNonQuery();
				}

				// Cập nhật trạng thái phiếu mượn
				if (modify.trasach(ID))
				{
					// Cập nhật thống kê trả sách
					modify.CapNhatThongKeTra(masinhvien, hoVaTen, soLuong, ngayTra);
					MessageBox.Show("Trả sách thành công", "Thông báo",
						MessageBoxButtons.OK, MessageBoxIcon.Information);
					LoadDataGridView();
				}
				else
				{
					// Nếu cập nhật trạng thái phiếu thất bại, khôi phục số lượng sách
					using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
					{
						sqlConnection.Open();
						string restoreQuery = "UPDATE quanlysach SET SoLuong = SoLuong - @SoLuongTra WHERE TenSach = @TenSach";
						SQLiteCommand restoreCommand = new SQLiteCommand(restoreQuery, sqlConnection);
						restoreCommand.Parameters.AddWithValue("@SoLuongTra", soLuong);
						restoreCommand.Parameters.AddWithValue("@TenSach", tenSach);
						restoreCommand.ExecuteNonQuery();
					}

					MessageBox.Show("Trả sách thất bại", "Lỗi",
						MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}




		private void button_kiemtra_Click(object sender, EventArgs e)
		{
			string timsv = this.textBox_timsinhvien.Text;
			using (SQLiteConnection sqliteConnection = Connection.GetSQLiteConnection())
			{
				sqliteConnection.Open();
				dataGridView1.DataSource = modify.timsv(timsv, sqliteConnection);
			}
		}

		private void textBox_timsinhvien_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				button_kiemtra_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}

		private void textBox_soluong_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				button_kiemtra_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}

		private void button1_Click(object sender, EventArgs e)
		{
			if (MessageBox.Show("Lịch hẹn sẽ hiện ở một cửa sổ riêng, bạn muốn mở chứ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				lichhenmuonsach lichhenmuonsach = new lichhenmuonsach();
				lichhenmuonsach.ShowDialog();
			}
		}

		private void button2_Click(object sender, EventArgs e)
		{
			if (MessageBox.Show("Bảng thống kê sinh viên quá hạn sẽ hiện ở một cửa sổ riêng, bạn muốn mở chứ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				thongkesinhvienquahan thongkesinhvienquahan = new thongkesinhvienquahan();
				thongkesinhvienquahan.ShowDialog();
			}
		}

		private void button_quetmamuonsach_Click(object sender, EventArgs e)
		{
			// Khi mở form quetmamuonsach
			quetmamuonsach quetForm = new quetmamuonsach();
			quetForm.Owner = this; // Đặt owner là muontrasach
			quetForm.ShowDialog();
		}

		private void button_quetmatrasach_Click(object sender, EventArgs e)
		{
			quetmatrasach quetForm = new quetmatrasach();
			quetForm.Owner = this; // Đặt owner là form muontrasach hiện tại
			quetForm.ShowDialog();
		}

		public void SetBookCode(string code)
		{
			textBox_masach.Text = code; // Đặt mã đã quét vào textbox
		}

		public void XuLyMaSachQuet(string maSach, string maSinhVien, string hoVaTen)
		{
			try
			{
				// Đảm bảo dữ liệu đã được tải lại để có thông tin mới nhất
				LoadDataGridView();

				// Tìm và chọn hàng có mã sách tương ứng
				bool timThay = false;
				foreach (DataGridViewRow row in dataGridView1.Rows)
				{
					if (row.Cells["MaSach"] != null &&
						row.Cells["MaSach"].Value != null &&
						row.Cells["MaSach"].Value.ToString() == maSach)
					{
						// Chọn hàng
						dataGridView1.ClearSelection();
						row.Selected = true;

						// Tìm một ô hiển thị để đặt làm CurrentCell
						DataGridViewCell visibleCell = null;
						foreach (DataGridViewCell cell in row.Cells)
						{
							if (cell.Visible && cell.OwningColumn.Visible)
							{
								visibleCell = cell;
								break;
							}
						}

						if (visibleCell != null)
						{
							dataGridView1.CurrentCell = visibleCell;
						}

						// Đảm bảo hàng được hiển thị trong vùng nhìn thấy
						if (dataGridView1.FirstDisplayedScrollingRowIndex > row.Index ||
							(dataGridView1.FirstDisplayedScrollingRowIndex + dataGridView1.DisplayedRowCount(true) - 1) < row.Index)
						{
							dataGridView1.FirstDisplayedScrollingRowIndex = row.Index;
						}

						timThay = true;
						break;
					}
				}

				// Đảm bảo UI được cập nhật
				Application.DoEvents();

				if (timThay)
				{
					// Hiển thị thông báo xác nhận
					if (MessageBox.Show($"Mã sinh viên: {maSinhVien}, Họ và tên: {hoVaTen} muốn trả sách?",
						"Xác nhận trả sách", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
					{
						// Gọi phương thức button_trasach_Click để trả sách
						button_trasach_Click(null, null);
					}
				}
				else
				{
					// Không tìm thấy mã sách trong DataGridView
					MessageBox.Show($"Không tìm thấy sách có mã {maSach} trong danh sách hiển thị.",
						"Không tìm thấy", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi xử lý mã sách: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

	}
}