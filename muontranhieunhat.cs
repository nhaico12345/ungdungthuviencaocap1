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
using DevExpress.XtraCharts;
using System.Data.SQLite;

namespace ungdungthuviencaocap
{
	public partial class muontranhieunhat : DevExpress.XtraEditors.XtraForm
	{
		private Modify modify;
		private bool isDetailView = false;

		public muontranhieunhat()
		{
			try
			{
				InitializeComponent();

				// Khởi tạo đối tượng Modify
				modify = new Modify();

				// Đặt giá trị mặc định cho combobox
				comboBox_thongkemuontra.SelectedIndex = 0;
				comboBox1.SelectedIndex = 0; // Mặc định là "Chế độ xem tổng hợp"

				// Đặt mặc định radioDay
				radioDay.Checked = true;
				dateTimePicker_chonngay.CustomFormat = "dd/MM/yyyy";

				// Tải thống kê ban đầu
				LoadStatistics();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khởi tạo form: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void comboBox_thongkemuontra_SelectedIndexChanged(object sender, EventArgs e)
		{
			try
			{
				isDetailView = false;
				LoadStatistics();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi thay đổi loại thống kê: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void dateTimePicker_chonngay_ValueChanged(object sender, EventArgs e)
		{
			try
			{
				LoadStatistics();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi thay đổi ngày: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void DateFilter_CheckedChanged(object sender, EventArgs e)
		{
			try
			{
				// Cập nhật định dạng ngày dựa trên lựa chọn lọc
				if (radioDay.Checked)
					dateTimePicker_chonngay.CustomFormat = "dd/MM/yyyy";
				else if (radioMonth.Checked)
					dateTimePicker_chonngay.CustomFormat = "MM/yyyy";
				else
					dateTimePicker_chonngay.CustomFormat = "yyyy";

				LoadStatistics();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi thay đổi bộ lọc ngày: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void LoadStatistics()
		{
			try
			{
				if (comboBox_thongkemuontra.SelectedItem == null)
					return;

				string selectedOption = comboBox_thongkemuontra.SelectedItem.ToString();
				DateTime selectedDate = dateTimePicker_chonngay.Value;

				// Lấy mẫu ngày cho bộ lọc
				string datePattern = GetDatePattern(selectedDate);

				// Lấy chế độ xem (tổng hợp/chi tiết)
				isDetailView = comboBox1.SelectedIndex == 1;

				// Lấy dữ liệu thống kê từ bảng ThongKeMuonTra
				DataTable statsData = modify.GetThongKeMuonTra(selectedOption, datePattern, isDetailView);
				dataGridView_muontrasach.DataSource = statsData;

				// Định dạng tiêu đề cột
				FormatDataGridView(selectedOption);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void FormatDataGridView(string selectedOption)
		{
			try
			{
				if (dataGridView_muontrasach.Columns.Count > 0)
				{
					dataGridView_muontrasach.Columns["MaSinhVien"].HeaderText = "Mã sinh viên";
					dataGridView_muontrasach.Columns["HoVaTen"].HeaderText = "Họ và tên";

					if (selectedOption == "Thống kê mượn sách")
					{
						dataGridView_muontrasach.Columns["SoLanMuon"].HeaderText = "Số lần mượn";
						dataGridView_muontrasach.Columns["TongSachMuon"].HeaderText = "Tổng số sách mượn";

						// Nếu ở chế độ xem chi tiết, định dạng cột ngày
						if (isDetailView && dataGridView_muontrasach.Columns.Contains("NgayMuon"))
						{
							dataGridView_muontrasach.Columns["NgayMuon"].HeaderText = "Ngày mượn";
						}
					}
					else // Thống kê trả sách
					{
						dataGridView_muontrasach.Columns["SoLanTra"].HeaderText = "Số lần trả";
						dataGridView_muontrasach.Columns["TongSachTra"].HeaderText = "Tổng số sách trả";

						// Nếu ở chế độ xem chi tiết, định dạng cột ngày
						if (isDetailView && dataGridView_muontrasach.Columns.Contains("NgayTra"))
						{
							dataGridView_muontrasach.Columns["NgayTra"].HeaderText = "Ngày trả";
						}
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi định dạng bảng dữ liệu: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private string GetDatePattern(DateTime date)
		{
			try
			{
				// Tạo mẫu ngày dựa trên bộ lọc đã chọn
				if (radioDay.Checked)
					return date.ToString("yyyy-MM-dd");
				else if (radioMonth.Checked)
					return date.ToString("yyyy-MM") + "%";
				else // Bộ lọc năm
					return date.ToString("yyyy") + "%";
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tạo mẫu ngày: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return "%"; // Trả về wildcard để lấy tất cả dữ liệu nếu có lỗi
			}
		}

		private void textBox_timkiem_KeyDown(object sender, KeyEventArgs e)
		{
			try
			{
				if (e.KeyCode == Keys.Enter)
				{
					TimKiemSinhVien();
					e.Handled = true;
					e.SuppressKeyPress = true;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi xử lý phím Enter: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void Button_timKiem_Click(object sender, EventArgs e)
		{
			try
			{
				TimKiemSinhVien();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tìm kiếm: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void TimKiemSinhVien()
		{
			try
			{
				if (comboBox_thongkemuontra.SelectedItem == null)
					return;

				string selectedOption = comboBox_thongkemuontra.SelectedItem.ToString();
				string maSinhVien = textBox_timkiem.Text.Trim();

				// Nếu ô tìm kiếm trống, hiển thị tất cả dữ liệu
				if (string.IsNullOrEmpty(maSinhVien))
				{
					LoadStatistics();
					return;
				}

				// Lấy chế độ xem
				isDetailView = comboBox1.SelectedIndex == 1;

				// Tìm kiếm theo mã sinh viên
				DataTable kqTimKiem = modify.TimKiemSinhVienThongKe(maSinhVien, selectedOption, isDetailView);
				dataGridView_muontrasach.DataSource = kqTimKiem;

				// Định dạng tiêu đề cột
				FormatDataGridView(selectedOption);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void Button_locngay_Click(object sender, EventArgs e)
		{
			try
			{
				// Gọi LoadStatistics để lọc theo ngày đã chọn
				LoadStatistics();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi lọc theo ngày: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
		{
			try
			{
				// Xác định xem chế độ xem chi tiết có được chọn không
				isDetailView = comboBox1.SelectedIndex == 1; // Chỉ số 1 là "Chế độ xem chi tiết"

				// Tải lại thống kê với chế độ xem mới
				LoadStatistics();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi thay đổi chế độ xem: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button_hientatcadulieu_Click(object sender, EventArgs e)
		{
			try
			{
				// Xác định chế độ xem (chi tiết hay tổng hợp)
				isDetailView = comboBox1.SelectedIndex == 1;

				// Lấy tất cả dữ liệu từ bảng ThongKeMuonTra
				DataTable allData = modify.GetAllThongKeMuonTra(isDetailView);

				// Hiển thị dữ liệu lên DataGridView
				dataGridView_muontrasach.DataSource = allData;

				// Định dạng hiển thị cho DataGridView
				FormatAllDataGridView();

				// Hiển thị thông báo trạng thái nếu có label_status
				if (Controls.Find("label_status", true).Length > 0)
				{
					Label statusLabel = (Label)Controls.Find("label_status", true)[0];
					statusLabel.Text = "Đang hiển thị: Tất cả dữ liệu";
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi hiển thị tất cả dữ liệu: " + ex.Message, "Lỗi",
								MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void FormatAllDataGridView()
		{
			try
			{
				if (dataGridView_muontrasach.Columns.Count > 0)
				{
					// Định dạng các cột chung
					dataGridView_muontrasach.Columns["MaSinhVien"].HeaderText = "Mã sinh viên";
					dataGridView_muontrasach.Columns["HoVaTen"].HeaderText = "Họ và tên";
					dataGridView_muontrasach.Columns["SoLanMuon"].HeaderText = "Số lần mượn";
					dataGridView_muontrasach.Columns["TongSachMuon"].HeaderText = "Tổng sách mượn";
					dataGridView_muontrasach.Columns["SoLanTra"].HeaderText = "Số lần trả";
					dataGridView_muontrasach.Columns["TongSachTra"].HeaderText = "Tổng sách trả";

					// Nếu ở chế độ xem chi tiết, định dạng thêm các cột ngày
					if (isDetailView)
					{
						if (dataGridView_muontrasach.Columns.Contains("NgayMuon"))
							dataGridView_muontrasach.Columns["NgayMuon"].HeaderText = "Ngày mượn";

						if (dataGridView_muontrasach.Columns.Contains("NgayTra"))
							dataGridView_muontrasach.Columns["NgayTra"].HeaderText = "Ngày trả";
					}

					// Tùy chỉnh chiều rộng cột nếu cần
					foreach (DataGridViewColumn column in dataGridView_muontrasach.Columns)
					{
						column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
					}

					// Làm nổi bật sinh viên mượn nhiều sách nhất
					if (dataGridView_muontrasach.Rows.Count > 0)
					{
						// Tìm giá trị lớn nhất trong cột SoLanMuon
						int maxIndex = 0;
						int maxValue = Convert.ToInt32(dataGridView_muontrasach.Rows[0].Cells["SoLanMuon"].Value);

						for (int i = 1; i < dataGridView_muontrasach.Rows.Count; i++)
						{
							int currentValue = Convert.ToInt32(dataGridView_muontrasach.Rows[i].Cells["SoLanMuon"].Value);
							if (currentValue > maxValue)
							{
								maxValue = currentValue;
								maxIndex = i;
							}
						}

						// Làm nổi bật hàng có giá trị lớn nhất
						dataGridView_muontrasach.Rows[maxIndex].DefaultCellStyle.BackColor = Color.LightGreen;
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi định dạng bảng dữ liệu tất cả: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
