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
	public partial class theloai : DevExpress.XtraEditors.XtraForm
	{
		Modify modify;
		private int selectedCategoryId = -1; // Biến để lưu ID của thể loại đang được chọn
		private string selectedCategoryName = null; // Biến để lưu tên thể loại đang được chọn cho việc lọc

		public theloai()
		{
			InitializeComponent();
			// Gán sự kiện CellClick cho dataGridView_theloai
			dataGridView_theloai.CellClick += dataGridView_theloai_CellClick;
			// Gán sự kiện Click cho button_sua
			button_sua.Click += button_sua_Click;
			// Gán sự kiện Click cho button_loc
			button_loc.Click += button_loc_Click;
			// Gán sự kiện Click cho button_lammoi (để hiển thị lại tất cả sách)
			button_lammoi.Click += button_lammoi_Click;
		}

		private void theloai_Load(object sender, EventArgs e)
		{
			try
			{
				modify = new Modify(); // Khởi tạo đối tượng Modify
				LoadCategoryDataGridView(); // Tải dữ liệu thể loại vào DataGridView
				LoadBookDataGridView(); // Tải dữ liệu sách ban đầu (tất cả sách)
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi tải form Thể loại: " + ex.Message, "Lỗi Khởi Tạo", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		// Phương thức này giờ sẽ tải tất cả sách hoặc sách đã lọc
		private void LoadBookDataGridView(DataTable bookData = null)
		{
			try
			{
				dataGridView_tacpham.DataSource = null; // Xóa nguồn dữ liệu cũ

				// Nếu không có dữ liệu sách cụ thể được cung cấp, tải tất cả sách
				if (bookData == null)
				{
					dataGridView_tacpham.DataSource = modify.getAllbooks();
				}
				else // Ngược lại, sử dụng dữ liệu sách đã được lọc
				{
					dataGridView_tacpham.DataSource = bookData;
				}


				if (dataGridView_tacpham.Columns.Count > 0)
				{
					// Cấu hình cột (ẩn ID, đặt HeaderText)
					try { dataGridView_tacpham.Columns["ID"].Visible = false; } catch { }
					try { dataGridView_tacpham.Columns["Masach"].HeaderText = "Mã sách"; } catch { }
					try { dataGridView_tacpham.Columns["TenSach"].HeaderText = "Tên sách"; } catch { }
					try { dataGridView_tacpham.Columns["TheLoai"].HeaderText = "Thể loại"; } catch { }
					try { dataGridView_tacpham.Columns["TacGia"].HeaderText = "Tác giả"; } catch { }
					try { dataGridView_tacpham.Columns["SoLuong"].HeaderText = "Số lượng"; } catch { }
					try { dataGridView_tacpham.Columns["NhaXuatBan"].HeaderText = "Nhà xuất bản"; } catch { }
					try { dataGridView_tacpham.Columns["NamXuatBan"].HeaderText = "Năm xuất bản"; } catch { }
					// Ẩn các cột không cần thiết nếu có
					if (dataGridView_tacpham.Columns.Contains("anh")) dataGridView_tacpham.Columns["anh"].Visible = false;
					if (dataGridView_tacpham.Columns.Contains("pdf")) dataGridView_tacpham.Columns["pdf"].Visible = false;
					if (dataGridView_tacpham.Columns.Contains("tomtatnoidung")) dataGridView_tacpham.Columns["tomtatnoidung"].Visible = false;


					// Cài đặt chung
					dataGridView_tacpham.ReadOnly = true;
					dataGridView_tacpham.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
					dataGridView_tacpham.AllowUserToAddRows = false;
					dataGridView_tacpham.MultiSelect = false; // Đổi thành false nếu chỉ muốn chọn 1 dòng sách
					dataGridView_tacpham.RowHeadersVisible = false;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi tải dữ liệu tác phẩm vào GridView: " + ex.Message, "Lỗi Tải Dữ Liệu", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void LoadCategoryDataGridView()
		{
			try
			{
				dataGridView_theloai.DataSource = null;
				DataTable categoryData = modify.GetAllTheLoai();
				dataGridView_theloai.DataSource = categoryData;

				if (dataGridView_theloai.Columns.Count > 0)
				{
					try { dataGridView_theloai.Columns["ID"].Visible = false; } catch { }
					try { dataGridView_theloai.Columns["MaTheLoai"].HeaderText = "Mã Thể Loại"; } catch { }
					try { dataGridView_theloai.Columns["TenTheLoai"].HeaderText = "Tên Thể Loại"; } catch { }
					try { dataGridView_theloai.Columns["MoTa"].HeaderText = "Mô Tả"; } catch { }
					try { dataGridView_theloai.Columns["TrangThai"].HeaderText = "Trạng Thái"; } catch { }
					try { dataGridView_theloai.Columns["MaTheLoai"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells; } catch { }
					try { dataGridView_theloai.Columns["TenTheLoai"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells; } catch { }
					try { dataGridView_theloai.Columns["TrangThai"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells; } catch { }
					try { dataGridView_theloai.Columns["MoTa"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill; } catch { }

					dataGridView_theloai.ReadOnly = true;
					dataGridView_theloai.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
					dataGridView_theloai.AllowUserToAddRows = false;
					dataGridView_theloai.AllowUserToDeleteRows = false;
					dataGridView_theloai.MultiSelect = false;
					dataGridView_theloai.RowHeadersVisible = false;
					dataGridView_theloai.AllowUserToResizeRows = false;
					dataGridView_theloai.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
					dataGridView_theloai.ColumnHeadersDefaultCellStyle.Font = new Font(DataGridView.DefaultFont, FontStyle.Bold);
					dataGridView_theloai.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
					dataGridView_theloai.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight;
					dataGridView_theloai.DefaultCellStyle.SelectionForeColor = SystemColors.HighlightText;
				}
			}
			catch (ArgumentException argEx)
			{
				MessageBox.Show($"Lỗi cấu hình cột DataGridView Thể loại: {argEx.Message}\nKiểm tra lại tên cột trong CSDL và câu lệnh SELECT.", "Lỗi Cấu Hình Grid", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi tải dữ liệu thể loại vào GridView: " + ex.Message, "Lỗi Tải Dữ Liệu", MessageBoxButtons.OK, MessageBoxIcon.Error);
				Console.WriteLine($"Error loading category data into GridView: {ex.StackTrace}");
			}
		}

		private void dataGridView_theloai_CellClick(object sender, DataGridViewCellEventArgs e)
		{
			if (e.RowIndex >= 0 && e.RowIndex < dataGridView_theloai.Rows.Count)
			{
				DataGridViewRow selectedRow = dataGridView_theloai.Rows[e.RowIndex];
				if (selectedRow.Cells["ID"].Value != null && int.TryParse(selectedRow.Cells["ID"].Value.ToString(), out int id))
				{
					selectedCategoryId = id;
				}
				else
				{
					selectedCategoryId = -1;
				}

				// Lưu tên thể loại được chọn
				selectedCategoryName = selectedRow.Cells["TenTheLoai"]?.Value?.ToString();

				textBox_matheloai.Text = selectedRow.Cells["MaTheLoai"]?.Value?.ToString() ?? "";
				textBox_tentheloai.Text = selectedCategoryName ?? ""; // Hiển thị tên đã lưu
				textBox_trangthai.Text = selectedRow.Cells["TrangThai"]?.Value?.ToString() ?? "";
				textBox_mota.Text = selectedRow.Cells["MoTa"]?.Value?.ToString() ?? "";
			}
		}

		private void button_sua_Click(object sender, EventArgs e)
		{
			string maTheLoai = textBox_matheloai.Text.Trim();
			string tenTheLoai = textBox_tentheloai.Text.Trim();
			string trangThai = textBox_trangthai.Text.Trim();
			string moTa = textBox_mota.Text.Trim();

			if (string.IsNullOrEmpty(maTheLoai))
			{
				MessageBox.Show("Mã thể loại không được để trống.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				textBox_matheloai.Focus();
				return;
			}
			if (string.IsNullOrEmpty(tenTheLoai))
			{
				MessageBox.Show("Tên thể loại không được để trống.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				textBox_tentheloai.Focus();
				return;
			}

			if (selectedCategoryId == -1) // Trường hợp thêm mới
			{
				DialogResult dialogResult = MessageBox.Show("Bạn muốn thêm Thể loại mới này vào dữ liệu?", "Xác Nhận Thêm Mới", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
				if (dialogResult == DialogResult.Yes)
				{
					try
					{
						bool result = modify.InsertTheLoai(maTheLoai, tenTheLoai, moTa, trangThai);
						if (result)
						{
							MessageBox.Show("Thêm thể loại mới thành công!", "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
							LoadCategoryDataGridView();
							ClearTextboxes();
						}
						else
						{
							MessageBox.Show("Thêm thể loại mới thất bại. Mã thể loại có thể đã tồn tại hoặc có lỗi xảy ra.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
					catch (Exception ex)
					{
						MessageBox.Show($"Lỗi khi thêm thể loại mới: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}
			else // Trường hợp cập nhật
			{
				try
				{
					bool result = modify.UpdateTheLoai(selectedCategoryId, maTheLoai, tenTheLoai, moTa, trangThai);
					if (result)
					{
						MessageBox.Show("Cập nhật thông tin thể loại thành công!", "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
						LoadCategoryDataGridView();
						// Nếu tên thể loại đã thay đổi, cập nhật selectedCategoryName và có thể lọc lại sách
						if (selectedCategoryName != tenTheLoai)
						{
							selectedCategoryName = tenTheLoai;
							// Tùy chọn: Tự động lọc lại sách với tên thể loại mới
							// FilterBooksByCategory(selectedCategoryName);
						}
						ClearTextboxes();
					}
					else
					{
						MessageBox.Show("Cập nhật thông tin thể loại thất bại. Mã thể loại có thể bị trùng với thể loại khác hoặc không có thay đổi nào được thực hiện.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi cập nhật thể loại: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}

		private void ClearTextboxes()
		{
			textBox_matheloai.Text = "";
			textBox_tentheloai.Text = "";
			textBox_trangthai.Text = "";
			textBox_mota.Text = "";
			selectedCategoryId = -1;
			selectedCategoryName = null; // Reset tên thể loại đã chọn
			dataGridView_theloai.ClearSelection();
			textBox_matheloai.Focus();
		}

		// Sự kiện click cho nút Lọc
		private void button_loc_Click(object sender, EventArgs e)
		{
			if (dataGridView_theloai.CurrentRow != null && !dataGridView_theloai.CurrentRow.IsNewRow)
			{
				// Lấy tên thể loại từ hàng đang được chọn (đã được lưu trong selectedCategoryName khi CellClick)
				// Hoặc lấy trực tiếp từ cell nếu muốn chắc chắn hơn:
				// string categoryNameToFilter = dataGridView_theloai.CurrentRow.Cells["TenTheLoai"]?.Value?.ToString();

				if (!string.IsNullOrEmpty(selectedCategoryName))
				{
					FilterBooksByCategory(selectedCategoryName);
				}
				else
				{
					MessageBox.Show("Không có tên thể loại hợp lệ để lọc.", "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					LoadBookDataGridView(); // Hiển thị lại tất cả sách nếu không có tên thể loại
				}
			}
			else
			{
				MessageBox.Show("Vui lòng chọn một thể loại từ danh sách để lọc.", "Chưa Chọn Thể Loại", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				LoadBookDataGridView(); // Hiển thị lại tất cả sách
			}
		}

		// Phương thức để lọc sách theo tên thể loại
		private void FilterBooksByCategory(string categoryName)
		{
			try
			{
				DataTable filteredBooks = modify.GetBooksByCategoryName(categoryName);
				LoadBookDataGridView(filteredBooks); // Tải dữ liệu sách đã lọc

				if (filteredBooks == null || filteredBooks.Rows.Count == 0)
				{
					MessageBox.Show($"Không tìm thấy sách nào thuộc thể loại '{categoryName}'.", "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi lọc sách theo thể loại: {ex.Message}", "Lỗi Lọc Dữ Liệu", MessageBoxButtons.OK, MessageBoxIcon.Error);
				LoadBookDataGridView(); // Tải lại tất cả sách nếu có lỗi
			}
		}

		// Sự kiện click cho nút Làm mới (hiển thị tất cả sách)
		private void button_lammoi_Click(object sender, EventArgs e)
		{
			LoadBookDataGridView(); // Tải lại tất cả sách
			ClearTextboxes(); // Xóa các trường nhập liệu và bỏ chọn thể loại
			dataGridView_theloai.ClearSelection(); // Đảm bảo không có thể loại nào được chọn
			selectedCategoryName = null; // Reset tên thể loại đã chọn
		}
	}
}
