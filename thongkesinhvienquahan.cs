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
	public partial class thongkesinhvienquahan : DevExpress.XtraEditors.XtraForm
	{
		private Modify modify;

		public thongkesinhvienquahan()
		{
			InitializeComponent();

			// Kết nối các sự kiện
			this.button_loadlaibang.Click += new EventHandler(this.button_loadlaibang_Click);
			this.button_xuatdanhsachsinhvienquahan.Click += new EventHandler(this.button_xuatdanhsachsinhvienquahan_Click);
			this.Load += new EventHandler(this.thongkesinhvienquahan_Load);

			modify = new Modify();
		}

		private void thongkesinhvienquahan_Load(object sender, EventArgs e)
		{
			LoadOverdueStudents();
			PopulateStudentComboBox();
		}

		private void LoadOverdueStudents()
		{
			try
			{
				// Lấy ngày hiện tại để so sánh với NgayTra
				DateTime currentDate = DateTime.Now.Date;
				string formattedDate = currentDate.ToString("yyyy-MM-dd");

				// Truy vấn cho sách quá hạn: IsActive = 1 (chưa trả) và NgayTra < ngày hiện tại
				string query = @"SELECT * FROM muontrasach 
                       WHERE IsActive = 1 AND date(NgayTra) < date(@CurrentDate)";

				using (SQLiteConnection con = Connection.GetSQLiteConnection())
				{
					con.Open();
					SQLiteCommand command = new SQLiteCommand(query, con);
					command.Parameters.AddWithValue("@CurrentDate", formattedDate);
					SQLiteDataAdapter adapter = new SQLiteDataAdapter(command);
					DataTable dt = new DataTable();
					adapter.Fill(dt);
					dataGridView1.DataSource = dt;

					// Ẩn cột ID và IsActive
					dataGridView1.Columns["ID"].Visible = false;
					dataGridView1.Columns["IsActive"].Visible = false;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void PopulateStudentComboBox()
		{
			try
			{
				if (dataGridView1.DataSource != null)
				{
					DataTable dt = (DataTable)dataGridView1.DataSource;
					// Sử dụng .ToString() để chuyển đổi giá trị sang chuỗi
					var distinctStudents = dt.AsEnumerable()
						.Select(row => row["MaSinhVien"].ToString())
						.Distinct()
						.ToList();

					comboBox_timmasinhvien.DataSource = distinctStudents;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}


		private void button_timmasinhvien_Click(object sender, EventArgs e)
		{
			try
			{
				string selectedStudentId = comboBox_timmasinhvien.Text;

				// Kiểm tra xem mã sinh viên có tồn tại trong dataGridView không
				if (dataGridView1.DataSource is DataTable dt)
				{
					DataView dv = dt.DefaultView;

					// Kiểm tra xem MaSinhVien có phải là số không
					if (long.TryParse(selectedStudentId, out _))
					{
						// Nếu là số, không cần dấu nháy đơn
						dv.RowFilter = $"MaSinhVien = {selectedStudentId}";
					}
					else
					{
						// Nếu là chuỗi, sử dụng dấu nháy đơn
						dv.RowFilter = $"MaSinhVien = '{selectedStudentId}'";
					}

					// Nếu không có hàng nào được tìm thấy sau khi lọc
					if (dv.Count == 0)
					{
						MessageBox.Show("Sinh viên này chưa mượn sách hoặc đã trả đúng thời hạn!",
							"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);

						// Tải lại dữ liệu gốc
						LoadOverdueStudents();
					}
					else
					{
						// Áp dụng bộ lọc để chỉ hiển thị sinh viên đã chọn
						dataGridView1.DataSource = dv.ToTable();
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}


		private void button_locngay_Click(object sender, EventArgs e)
		{
			try
			{
				DateTime selectedDate = dateTimePicker1.Value.Date;
				string formattedDate = selectedDate.ToString("yyyy-MM-dd");

				// Lọc để hiển thị sách quá hạn có ngày trả bằng ngày đã chọn
				string query = @"SELECT * FROM muontrasach 
                               WHERE IsActive = 1 AND date(NgayTra) < date(@CurrentDate) 
                               AND date(NgayTra) = date(@SelectedDate)";

				using (SQLiteConnection con = Connection.GetSQLiteConnection())
				{
					con.Open();
					SQLiteCommand command = new SQLiteCommand(query, con);
					command.Parameters.AddWithValue("@CurrentDate", DateTime.Now.ToString("yyyy-MM-dd"));
					command.Parameters.AddWithValue("@SelectedDate", formattedDate);
					SQLiteDataAdapter adapter = new SQLiteDataAdapter(command);
					DataTable dt = new DataTable();
					adapter.Fill(dt);

					if (dt.Rows.Count == 0)
					{
						MessageBox.Show("Không có sinh viên nào quá hạn trả sách vào ngày " +
							selectedDate.ToString("dd/MM/yyyy"), "Thông báo",
							MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					else
					{
						dataGridView1.DataSource = dt;
						PopulateStudentComboBox(); // Cập nhật combobox với dữ liệu đã lọc
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button_loadlaibang_Click(object sender, EventArgs e)
		{
			LoadOverdueStudents();
			PopulateStudentComboBox();
		}

		private void button_xuatdanhsachsinhvienquahan_Click(object sender, EventArgs e)
		{
			try
			{
				if (dataGridView1.DataSource != null && dataGridView1.Rows.Count > 0)
				{
					SaveFileDialog saveFileDialog = new SaveFileDialog();
					saveFileDialog.Filter = "Excel Files|*.xlsx";
					saveFileDialog.Title = "Xuất danh sách sinh viên quá hạn";
					saveFileDialog.FileName = "DanhSachSinhVienQuaHan_" + DateTime.Now.ToString("yyyyMMdd");

					if (saveFileDialog.ShowDialog() == DialogResult.OK)
					{
						DataTable dt = (DataTable)dataGridView1.DataSource;
						string filePath = saveFileDialog.FileName;

						// Xuất sang Excel sử dụng phương thức có sẵn trong lớp Modify
						using (var workbook = new ClosedXML.Excel.XLWorkbook())
						{
							var worksheet = workbook.Worksheets.Add(dt, "SinhVienQuaHan");
							workbook.SaveAs(filePath);
						}

						MessageBox.Show("Xuất danh sách thành công!", "Thông báo",
							MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				}
				else
				{
					MessageBox.Show("Không có dữ liệu để xuất!", "Thông báo",
						MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi xuất danh sách: " + ex.Message, "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
