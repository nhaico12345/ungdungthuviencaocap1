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
using System.IO; // Namespace cho thao tác File, Directory
using System.Diagnostics; // Namespace cho Process (mở file)

namespace ungdungthuviencaocap
{
	public partial class timkiemsach : DevExpress.XtraEditors.XtraForm
	{
		Modify modify;
		private int selectedBookId = -1; // Lưu ID sách đang chọn
		private string selectedBookMasach = null; // Lưu mã sách đang chọn

		// Đường dẫn thư mục gốc cho dữ liệu QLTV trong Documents
		private string baseDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "QLTV");
		private string pdfFolder;
		private string imageFolder;

		public timkiemsach()
		{
			InitializeComponent();

			// Khởi tạo đường dẫn thư mục
			pdfFolder = Path.Combine(baseDataFolder, "pdf");
			imageFolder = Path.Combine(baseDataFolder, "anhsach");

			// Đảm bảo các thư mục tồn tại khi form load
			EnsureDirectoriesExist();

			// Xóa các label và vô hiệu hóa nút khi khởi động
			ClearLabelsAndSelection();
			DisableActionButtons();

			// Đăng ký sự kiện
			this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellClick);
			this.button_capnhatanh.Click += new System.EventHandler(this.button_capnhatanh_Click);
			this.button_capnhatfilepdf.Click += new System.EventHandler(this.button_capnhatfilepdf_Click);
			this.button_mofilepdf.Click += new System.EventHandler(this.button_mofilepdf_Click);
		}

		/// <summary>
		/// Đảm bảo các thư mục lưu trữ PDF và ảnh tồn tại.
		/// </summary>
		private void EnsureDirectoriesExist()
		{
			try
			{
				if (!Directory.Exists(pdfFolder))
				{
					Directory.CreateDirectory(pdfFolder);
				}
				if (!Directory.Exists(imageFolder))
				{
					Directory.CreateDirectory(imageFolder);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Không thể tạo thư mục lưu trữ cần thiết ({baseDataFolder}): {ex.Message}\nVui lòng kiểm tra quyền truy cập.",
								"Lỗi tạo thư mục", MessageBoxButtons.OK, MessageBoxIcon.Error);
				// Consider disabling functionality if directories cannot be created
			}
		}

		/// <summary>
		/// Xóa nội dung các label và đặt lại ID sách đã chọn.
		/// </summary>
		private void ClearLabelsAndSelection()
		{
			label_tensach.Text = "";
			label_tacgia.Text = "";
			label_theloai.Text = "";
			label_nhaxuatban.Text = "";
			label_namxuatban.Text = "";
			label_soluong.Text = "";
			pictureBox3.Image = null; // Clear image
			selectedBookId = -1; // Reset ID
			selectedBookMasach = null; // Reset Masach
		}

		/// <summary>
		/// Vô hiệu hóa các nút hành động (Cập nhật, Mở file).
		/// </summary>
		private void DisableActionButtons()
		{
			button_capnhatanh.Enabled = false;
			button_capnhatfilepdf.Enabled = false;
			button_mofilepdf.Enabled = false;
		}

		/// <summary>
		/// Kích hoạt các nút hành động.
		/// </summary>
		private void EnableActionButtons()
		{
			button_capnhatanh.Enabled = true;
			button_capnhatfilepdf.Enabled = true;
			button_mofilepdf.Enabled = true;
		}

		/// <summary>
		/// Tải và cấu hình dữ liệu cho DataGridView.
		/// </summary>
		private void LoadDataGridView()
		{
			try
			{
				dataGridView1.DataSource = null;
				modify = modify ?? new Modify();
				// *** SỬA Ở ĐÂY: Gọi đúng tên phương thức getAllbooks ***
				DataTable dt = modify.getAllbooks(); // Sử dụng getAllbooks()
				dataGridView1.DataSource = dt;

				// Configure columns
				if (dataGridView1.Columns.Count > 0)
				{
					// Hide unnecessary columns
					string[] columnsToHide = { "ID", "Masach", "anh", "pdf", "tomtatnoidung" };
					foreach (string colName in columnsToHide)
					{
						if (dataGridView1.Columns.Contains(colName))
						{
							dataGridView1.Columns[colName].Visible = false;
						}
					}

					// Set HeaderText and display order (example)
					Dictionary<string, Tuple<string, int>> columnConfig = new Dictionary<string, Tuple<string, int>>
					{
						{ "TenSach", Tuple.Create("Tên sách", 0) },
						{ "TheLoai", Tuple.Create("Thể loại", 1) },
						{ "TacGia", Tuple.Create("Tác giả", 2) },
						{ "NhaXuatBan", Tuple.Create("Nhà xuất bản", 3) },
						{ "NamXuatBan", Tuple.Create("Năm xuất bản", 4) },
						{ "SoLuong", Tuple.Create("Số lượng", 5) }
					};

					foreach (var config in columnConfig)
					{
						if (dataGridView1.Columns.Contains(config.Key))
						{
							dataGridView1.Columns[config.Key].HeaderText = config.Value.Item1;
							dataGridView1.Columns[config.Key].DisplayIndex = config.Value.Item2;
						}
					}

					// General settings
					dataGridView1.ReadOnly = true;
					dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
					dataGridView1.AllowUserToAddRows = false;
					dataGridView1.MultiSelect = false;
					dataGridView1.RowHeadersVisible = false;
					dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
					dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font(DataGridView.DefaultFont, FontStyle.Bold);
					dataGridView1.ColumnHeadersVisible = false; // Headers still hidden
					dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
				}
			}
			catch (ArgumentException argEx) // Column name error
			{
				MessageBox.Show($"Lỗi cấu hình cột DataGridView: {argEx.Message}\nKiểm tra lại tên cột.", "Lỗi Cấu Hình Grid", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			catch (Exception ex) // General loading error
			{
				MessageBox.Show("Lỗi khi tải dữ liệu vào GridView: " + ex.Message, "Lỗi Tải Dữ Liệu", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		/// <summary>
		/// Handles the Load event of the form.
		/// </summary>
		private void timkiemsach_Load(object sender, EventArgs e)
		{
			try
			{
				modify = new Modify();
				LoadDataGridView(); // Calls the method that now uses getAllbooks()
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi tải form Tìm kiếm sách: " + ex.Message, "Lỗi Khởi Tạo", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		/// <summary>
		/// Handles the CellClick event of the DataGridView.
		/// </summary>
		private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
		{
			// Validate click
			if (e.RowIndex >= 0 && e.RowIndex < dataGridView1.Rows.Count && !dataGridView1.Rows[e.RowIndex].IsNewRow)
			{
				try
				{
					DataGridViewRow row = dataGridView1.Rows[e.RowIndex];

					// Safely get ID and Masach
					selectedBookId = -1; // Reset before trying
					if (row.Cells["ID"]?.Value != null && int.TryParse(row.Cells["ID"].Value.ToString(), out int id))
					{
						selectedBookId = id;
					}
					selectedBookMasach = row.Cells["Masach"]?.Value?.ToString(); // Get Masach

					// Display info on Labels (handle null)
					label_tensach.Text = row.Cells["TenSach"]?.Value?.ToString() ?? "N/A";
					label_tacgia.Text = row.Cells["TacGia"]?.Value?.ToString() ?? "N/A";
					label_theloai.Text = row.Cells["TheLoai"]?.Value?.ToString() ?? "N/A";
					label_nhaxuatban.Text = row.Cells["NhaXuatBan"]?.Value?.ToString() ?? "N/A";
					label_namxuatban.Text = row.Cells["NamXuatBan"]?.Value?.ToString() ?? "N/A";
					label_soluong.Text = row.Cells["SoLuong"]?.Value?.ToString() ?? "N/A";

					// Display image
					string imagePath = row.Cells["anh"]?.Value?.ToString();
					pictureBox3.Image = null; // Clear old image first
					if (!string.IsNullOrEmpty(imagePath) && File.Exists(imagePath))
					{
						try
						{
							// Load image safely, without locking the file
							using (var stream = new FileStream(imagePath, FileMode.Open, FileAccess.Read))
							{
								pictureBox3.Image = Image.FromStream(stream);
								pictureBox3.SizeMode = PictureBoxSizeMode.Zoom; // Or StretchImage
							}
						}
						catch (Exception imgEx)
						{
							Console.WriteLine($"Error loading image '{imagePath}': {imgEx.Message}");
							// Avoid showing MessageBox here to prevent user annoyance
							pictureBox3.Image = null; // Ensure image is cleared if load fails
						}
					}

					// Enable/Disable buttons based on valid ID selection
					if (selectedBookId > 0)
					{
						EnableActionButtons();
					}
					else
					{
						DisableActionButtons();
					}
				}
				catch (Exception ex) // Catch general errors during click processing
				{
					MessageBox.Show("Lỗi khi hiển thị chi tiết sách: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					ClearLabelsAndSelection(); // Clear info
					DisableActionButtons(); // Disable buttons
				}
			}
			else // Clicked on header or empty area
			{
				ClearLabelsAndSelection();
				DisableActionButtons();
			}
		}

		/// <summary>
		/// Handles the Click event for the Update PDF button.
		/// </summary>
		private void button_capnhatfilepdf_Click(object sender, EventArgs e)
		{
			if (selectedBookId <= 0) // Check if a book is selected
			{
				MessageBox.Show("Vui lòng chọn một sách từ danh sách trước.", "Chưa chọn sách", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			try
			{
				using (OpenFileDialog openFileDialog = new OpenFileDialog())
				{
					openFileDialog.Filter = "PDF Files (*.pdf)|*.pdf|All files (*.*)|*.*";
					openFileDialog.Title = "Chọn file PDF cho sách";
					openFileDialog.CheckFileExists = true;
					openFileDialog.CheckPathExists = true;

					if (openFileDialog.ShowDialog() == DialogResult.OK)
					{
						string selectedFilePath = openFileDialog.FileName;
						string targetFileName = Path.GetFileName(selectedFilePath);
						// Create unique target filename using Masach (if available)
						if (!string.IsNullOrEmpty(selectedBookMasach))
						{
							string cleanMasach = string.Join("_", selectedBookMasach.Split(Path.GetInvalidFileNameChars()));
							targetFileName = $"{cleanMasach}_{targetFileName}";
						}
						string targetFilePath = Path.Combine(pdfFolder, targetFileName);

						// Copy file (overwrite if exists)
						try
						{
							File.Copy(selectedFilePath, targetFilePath, true);
						}
						catch (IOException ioEx) // Specific IO error
						{
							MessageBox.Show($"Lỗi khi sao chép file PDF: {ioEx.Message}\nĐảm bảo bạn có quyền ghi vào thư mục:\n{pdfFolder}", "Lỗi sao chép file", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return; // Stop if copy fails
						}
						catch (Exception exCopy) // Other copy errors
						{
							MessageBox.Show($"Lỗi không xác định khi sao chép file PDF: {exCopy.Message}", "Lỗi sao chép file", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return; // Stop
						}

						// Update database path
						if (modify.UpdateBookPdfPath(selectedBookId, targetFilePath))
						{
							MessageBox.Show("Cập nhật file PDF thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);
							// Optionally reload grid if pdf column is visible
							// LoadDataGridView();
						}
						else
						{
							// DB Error message shown by Modify
							// Consider deleting the copied file if DB update failed
							try { if (File.Exists(targetFilePath)) File.Delete(targetFilePath); } catch (Exception exDel) { Console.WriteLine($"Error deleting temp PDF: {exDel.Message}"); }
						}
					}
				}
			}
			catch (Exception ex) // Catch general process errors
			{
				MessageBox.Show($"Đã xảy ra lỗi không mong muốn khi cập nhật PDF: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		/// <summary>
		/// Handles the Click event for the Update Image button.
		/// </summary>
		private void button_capnhatanh_Click(object sender, EventArgs e)
		{
			if (selectedBookId <= 0) // Check book selection
			{
				MessageBox.Show("Vui lòng chọn một sách từ danh sách trước.", "Chưa chọn sách", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			try
			{
				using (OpenFileDialog openFileDialog = new OpenFileDialog())
				{
					openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.bmp|All files (*.*)|*.*";
					openFileDialog.Title = "Chọn file ảnh cho sách";
					openFileDialog.CheckFileExists = true;
					openFileDialog.CheckPathExists = true;

					if (openFileDialog.ShowDialog() == DialogResult.OK)
					{
						string selectedFilePath = openFileDialog.FileName;
						string targetFileName = Path.GetFileName(selectedFilePath);
						// Create unique target filename using Masach
						if (!string.IsNullOrEmpty(selectedBookMasach))
						{
							string cleanMasach = string.Join("_", selectedBookMasach.Split(Path.GetInvalidFileNameChars()));
							targetFileName = $"{cleanMasach}_{targetFileName}";
						}
						string targetFilePath = Path.Combine(imageFolder, targetFileName);

						// Copy file (overwrite)
						try
						{
							File.Copy(selectedFilePath, targetFilePath, true);
						}
						catch (IOException ioEx)
						{
							MessageBox.Show($"Lỗi khi sao chép file ảnh: {ioEx.Message}\nĐảm bảo bạn có quyền ghi vào thư mục:\n{imageFolder}", "Lỗi sao chép file", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return;
						}
						catch (Exception exCopy)
						{
							MessageBox.Show($"Lỗi không xác định khi sao chép file ảnh: {exCopy.Message}", "Lỗi sao chép file", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return;
						}

						// Update Database
						if (modify.UpdateBookImagePath(selectedBookId, targetFilePath))
						{
							MessageBox.Show("Cập nhật ảnh thành công!", "Thành công", MessageBoxButtons.OK, MessageBoxIcon.Information);

							// Display the new image in PictureBox
							pictureBox3.Image = null; // Clear old image
							if (File.Exists(targetFilePath))
							{
								try
								{
									using (var stream = new FileStream(targetFilePath, FileMode.Open, FileAccess.Read))
									{
										pictureBox3.Image = Image.FromStream(stream);
										pictureBox3.SizeMode = PictureBoxSizeMode.Zoom;
									}
								}
								catch (Exception imgEx) { Console.WriteLine($"Error loading new image: {imgEx.Message}"); }
							}
							// Optionally update grid
							// LoadDataGridView();
						}
						else
						{
							// DB Error reported by Modify
							try { if (File.Exists(targetFilePath)) File.Delete(targetFilePath); } catch (Exception exDel) { Console.WriteLine($"Error deleting temp image: {exDel.Message}"); }
						}
					}
				}
			}
			catch (Exception ex) // General errors
			{
				MessageBox.Show($"Đã xảy ra lỗi không mong muốn khi cập nhật ảnh: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		/// <summary>
		/// Handles the Click event for the Open PDF button.
		/// </summary>
		private void button_mofilepdf_Click(object sender, EventArgs e)
		{
			if (selectedBookId <= 0) // Check book selection
			{
				MessageBox.Show("Vui lòng chọn một sách từ danh sách trước.", "Chưa chọn sách", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			try
			{
				// Get path from DB
				string pdfPath = modify.GetBookPdfPath(selectedBookId);

				// Check path and file existence
				if (string.IsNullOrEmpty(pdfPath))
				{
					MessageBox.Show("File PDF chưa được thêm cho sách này. Vui lòng cập nhật.", "Chưa có file", MessageBoxButtons.OK, MessageBoxIcon.Information); // Corrected icon
					return;
				}
				if (!File.Exists(pdfPath))
				{
					MessageBox.Show($"Không tìm thấy file PDF tại đường dẫn đã lưu:\n{pdfPath}\n\nFile có thể đã bị xóa hoặc di chuyển. Vui lòng cập nhật lại.", "Không tìm thấy file", MessageBoxButtons.OK, MessageBoxIcon.Warning); // Corrected icon
					return;
				}

				// Open file using default application
				try
				{
					ProcessStartInfo psi = new ProcessStartInfo
					{
						FileName = pdfPath,
						UseShellExecute = true // Important for using default app
					};
					Process.Start(psi);
				}
				catch (Win32Exception winEx) // No PDF reader installed
				{
					MessageBox.Show($"Không thể mở file PDF. Lỗi: {winEx.Message}\nĐảm bảo bạn đã cài đặt trình đọc PDF.", "Lỗi mở file", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				catch (Exception exOpen) // Other errors during opening
				{
					MessageBox.Show($"Đã xảy ra lỗi khi mở file PDF: {exOpen.Message}", "Lỗi mở file", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			catch (Exception ex) // General errors
			{
				MessageBox.Show($"Đã xảy ra lỗi không mong muốn khi mở PDF: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button_locthongtin_Click(object sender, EventArgs e)
		{
			try
			{
				if (MessageBox.Show("Sẽ hiện cửa sổ lọc thông tin vẫn tiếp tục chứ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
				{
					locthongtin locthongtin = new locthongtin();
					locthongtin.ShowDialog();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở bảng thống kê: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		/// <summary>
		/// Cập nhật DataGridView với kết quả đã lọc
		/// </summary>
		public void CapNhatDataGridVoiKetQuaLoc(DataTable duLieuDaLoc)
		{
			try
			{
				dataGridView1.DataSource = duLieuDaLoc;
				CauHinhCotDataGridView();
				ClearLabelsAndSelection(); // Xóa panel chi tiết khi thay đổi dữ liệu
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi cập nhật dữ liệu đã lọc: {ex.Message}",
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		/// <summary>
		/// Tải lại dữ liệu gốc không có bộ lọc
		/// </summary>
		public void TaiLaiDuLieuGoc()
		{
			try
			{
				LoadDataGridView();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tải lại dữ liệu gốc: {ex.Message}",
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		/// <summary>
		/// Cấu hình các cột của DataGridView sau khi nguồn dữ liệu thay đổi
		/// </summary>
		private void CauHinhCotDataGridView()
		{
			try
			{
				if (dataGridView1.Columns.Count > 0)
				{
					// Ẩn các cột không cần thiết
					string[] cotAnDi = { "ID", "Masach", "anh", "pdf", "tomtatnoidung" };
					foreach (string tenCot in cotAnDi)
					{
						if (dataGridView1.Columns.Contains(tenCot))
						{
							dataGridView1.Columns[tenCot].Visible = false;
						}
					}

					// Đặt tiêu đề và thứ tự hiển thị
					Dictionary<string, Tuple<string, int>> cauHinhCot = new Dictionary<string, Tuple<string, int>>
			{
				{ "TenSach", Tuple.Create("Tên sách", 0) },
				{ "TheLoai", Tuple.Create("Thể loại", 1) },
				{ "TacGia", Tuple.Create("Tác giả", 2) },
				{ "NhaXuatBan", Tuple.Create("Nhà xuất bản", 3) },
				{ "NamXuatBan", Tuple.Create("Năm xuất bản", 4) },
				{ "SoLuong", Tuple.Create("Số lượng", 5) }
			};

					foreach (var cauHinh in cauHinhCot)
					{
						if (dataGridView1.Columns.Contains(cauHinh.Key))
						{
							dataGridView1.Columns[cauHinh.Key].HeaderText = cauHinh.Value.Item1;
							dataGridView1.Columns[cauHinh.Key].DisplayIndex = cauHinh.Value.Item2;
						}
					}

					// Áp dụng các thiết lập chung
					dataGridView1.ReadOnly = true;
					dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
					dataGridView1.AllowUserToAddRows = false;
					dataGridView1.MultiSelect = false;
					dataGridView1.RowHeadersVisible = false;
					dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
					dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font(DataGridView.DefaultFont, FontStyle.Bold);
					dataGridView1.ColumnHeadersVisible = false;
					dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi cấu hình cột DataGridView: {ex.Message}",
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button_loadlaiform_Click(object sender, EventArgs e)
		{
			timkiemsach_Load(sender, e); // Gọi lại hàm load form
		}
	}
}