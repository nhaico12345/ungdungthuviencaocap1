// tacgia.cs
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms; // <- Namespace này chứa ComboBox chuẩn
using DevExpress.XtraEditors; // <- Namespace này cũng có thể có ComboBox hoặc ComboBoxEdit
using System.IO; // Cần thiết cho Path, Directory, File

namespace ungdungthuviencaocap
{
	public partial class tacgia : DevExpress.XtraEditors.XtraForm
	{
		Modify modify;
		private int selectedAuthorId = -1; // Lưu ID của tác giả đang được chọn
		private string selectedAuthorMaTacGia = null; // Lưu Mã Tác Giả để tạo tên file ảnh
		private string currentImagePath = null; // Lưu đường dẫn ảnh hiện tại của tác giả đang chọn

		public tacgia()
		{
			InitializeComponent();
			// Thiết lập ban đầu cho DateTimePicker (nếu muốn hiển thị trống)
			dateTimePicker_ngaysinh.Format = DateTimePickerFormat.Custom;
			dateTimePicker_ngaysinh.CustomFormat = " "; // Hiển thị trống ban đầu
			dateTimePicker_ngaysinh.Checked = false; // Bỏ check ban đầu
		}

		private void tacgia_Load(object sender, EventArgs e)
		{
			try
			{
				modify = new Modify(); // Khởi tạo đối tượng Modify
				LoadAuthorDataGridView(); // Tải dữ liệu tác giả vào DataGridView
				LoadBookDataGridView(); // Tải dữ liệu sách ban đầu
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi tải form Tác giả: " + ex.Message, "Lỗi Khởi Tạo", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		/// <summary>
		/// Tải dữ liệu từ bảng tacgia vào dataGridView_tacgia và cấu hình.
		/// </summary>
		private void LoadAuthorDataGridView()
		{
			try
			{
				dataGridView_tacgia.DataSource = null;
				DataTable authorData = modify.GetAllAuthors(); // Giả sử hàm này đã tồn tại trong Modify.cs
				dataGridView_tacgia.DataSource = authorData;


				if (dataGridView_tacgia.Columns.Count > 0)
				{
					dataGridView_tacgia.Columns["ID"].Visible = false;
					dataGridView_tacgia.Columns["HocHam"].Visible = false;
					dataGridView_tacgia.Columns["HocVi"].Visible = false;
					dataGridView_tacgia.Columns["DiaChi"].Visible = false;
					dataGridView_tacgia.Columns["Email"].Visible = false;
					dataGridView_tacgia.Columns["QuocTich"].Visible = false;
					dataGridView_tacgia.Columns["Anh"].Visible = false;
					dataGridView_tacgia.Columns["TieuSu"].Visible = false;
					dataGridView_tacgia.Columns["TrangThai"].Visible = false;

					// --- Cấu hình cột (Header Text, Visible, Format, AutoSizeMode) ---
					dataGridView_tacgia.Columns["ID"].HeaderText = "ID";
					dataGridView_tacgia.Columns["MaTacGia"].HeaderText = "Mã Tác Giả";
					dataGridView_tacgia.Columns["HoTen"].HeaderText = "Họ Tên";
					dataGridView_tacgia.Columns["HocHam"].HeaderText = "Học Hàm";
					dataGridView_tacgia.Columns["HocVi"].HeaderText = "Học Vị";
					dataGridView_tacgia.Columns["DiaChi"].HeaderText = "Địa Chỉ";
					dataGridView_tacgia.Columns["Email"].HeaderText = "Email";
					dataGridView_tacgia.Columns["NgaySinh"].HeaderText = "Ngày Sinh";
					dataGridView_tacgia.Columns["QuocTich"].HeaderText = "Quốc Tịch";
					dataGridView_tacgia.Columns["Anh"].HeaderText = "Ảnh";
					dataGridView_tacgia.Columns["TieuSu"].HeaderText = "Tiểu Sử";
					dataGridView_tacgia.Columns["TrangThai"].HeaderText = "Trạng Thái";

					dataGridView_tacgia.Columns["ID"].Visible = false;
					dataGridView_tacgia.Columns["Anh"].Visible = false;
					dataGridView_tacgia.Columns["TieuSu"].Visible = false;
					dataGridView_tacgia.Columns["DiaChi"].Visible = false;

					try
					{
						dataGridView_tacgia.Columns["NgaySinh"].DefaultCellStyle.Format = "dd/MM/yyyy";
						dataGridView_tacgia.Columns["NgaySinh"].DefaultCellStyle.NullValue = "";
					}
					catch { /* Ignore formatting errors */ }

					dataGridView_tacgia.Columns["MaTacGia"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
					dataGridView_tacgia.Columns["HoTen"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
					dataGridView_tacgia.Columns["HocHam"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
					dataGridView_tacgia.Columns["HocVi"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
					dataGridView_tacgia.Columns["Email"].MinimumWidth = 120;
					dataGridView_tacgia.Columns["Email"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
					dataGridView_tacgia.Columns["TrangThai"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

					// --- Cài đặt chung cho DataGridView ---
					dataGridView_tacgia.ReadOnly = true;
					dataGridView_tacgia.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
					dataGridView_tacgia.AllowUserToAddRows = false;
					dataGridView_tacgia.AllowUserToDeleteRows = false;
					dataGridView_tacgia.MultiSelect = false;
					dataGridView_tacgia.RowHeadersVisible = false;
					dataGridView_tacgia.AllowUserToResizeRows = false;
					dataGridView_tacgia.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(240, 240, 240);
					dataGridView_tacgia.ColumnHeadersDefaultCellStyle.Font = new Font(DataGridView.DefaultFont, FontStyle.Bold);
					dataGridView_tacgia.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
					dataGridView_tacgia.DefaultCellStyle.SelectionBackColor = SystemColors.Highlight;
					dataGridView_tacgia.DefaultCellStyle.SelectionForeColor = SystemColors.HighlightText;

					// Gán sự kiện CellClick (đảm bảo không bị gán nhiều lần nếu có)
					dataGridView_tacgia.CellClick -= dataGridView_tacgia_CellClick; // Xóa handler cũ (nếu có)
					dataGridView_tacgia.CellClick += dataGridView_tacgia_CellClick; // Gán handler mới

				}
			}
			catch (ArgumentException argEx)
			{
				MessageBox.Show($"Lỗi cấu hình cột DataGridView Tác giả: {argEx.Message}\nKiểm tra lại tên cột trong CSDL và câu lệnh SELECT.", "Lỗi Cấu Hình Grid", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi tải dữ liệu tác giả vào GridView: " + ex.Message, "Lỗi Tải Dữ Liệu", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		/// <summary>
		/// Tải dữ liệu từ bảng quanlysach vào dataGridView_tacpham và cấu hình.
		/// </summary>
		private void LoadBookDataGridView()
		{
			try
			{
				dataGridView_tacpham.DataSource = null;
				dataGridView_tacpham.DataSource = modify.getAllbooks(); // Sử dụng hàm lấy sách hiện có

				if (dataGridView_tacpham.Columns.Count > 0)
				{
					// Cấu hình cột (ẩn ID, đặt HeaderText)
					try { dataGridView_tacpham.Columns["ID"].Visible = false; } catch { }
					try { dataGridView_tacpham.Columns["anh"].Visible = false; } catch { }
					try { dataGridView_tacpham.Columns["pdf"].Visible = false; } catch { }
					try { dataGridView_tacpham.Columns["tomtatnoidung"].Visible = false; } catch { }
					try { dataGridView_tacpham.Columns["Masach"].HeaderText = "Mã sách"; } catch { }
					try { dataGridView_tacpham.Columns["TenSach"].HeaderText = "Tên sách"; } catch { }
					try { dataGridView_tacpham.Columns["TheLoai"].HeaderText = "Thể loại"; } catch { }
					try { dataGridView_tacpham.Columns["TacGia"].HeaderText = "Tác giả"; } catch { }
					try { dataGridView_tacpham.Columns["SoLuong"].HeaderText = "Số lượng"; } catch { }
					try { dataGridView_tacpham.Columns["NhaXuatBan"].HeaderText = "Nhà xuất bản"; } catch { }
					try { dataGridView_tacpham.Columns["NamXuatBan"].HeaderText = "Năm xuất bản"; } catch { }

					// Cài đặt chung
					dataGridView_tacpham.ReadOnly = true;
					dataGridView_tacpham.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
					dataGridView_tacpham.AllowUserToAddRows = false;
					dataGridView_tacpham.MultiSelect = false;
					dataGridView_tacpham.RowHeadersVisible = false;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi tải dữ liệu tác phẩm vào GridView: " + ex.Message, "Lỗi Tải Dữ Liệu", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void dataGridView_tacgia_CellClick(object sender, DataGridViewCellEventArgs e)
		{
			if (e.RowIndex >= 0 && e.RowIndex < dataGridView_tacgia.Rows.Count)
			{
				DataGridViewRow selectedRow = dataGridView_tacgia.Rows[e.RowIndex];
				try
				{
					selectedAuthorId = Convert.ToInt32(selectedRow.Cells["ID"].Value);
					selectedAuthorMaTacGia = selectedRow.Cells["MaTacGia"]?.Value?.ToString();
					currentImagePath = selectedRow.Cells["Anh"]?.Value?.ToString();

					textBox_hovaten.Text = selectedRow.Cells["HoTen"]?.Value?.ToString() ?? "";

					// Xử lý Ngày Sinh
					object ngaySinhValue = selectedRow.Cells["NgaySinh"]?.Value;
					if (ngaySinhValue != null && ngaySinhValue != DBNull.Value)
					{
						string dateString = ngaySinhValue.ToString();
						if (DateTime.TryParse(dateString, out DateTime dob) && dob <= DateTime.Today)
						{
							dateTimePicker_ngaysinh.Value = dob;
							dateTimePicker_ngaysinh.Checked = true;
							dateTimePicker_ngaysinh.Format = DateTimePickerFormat.Custom;
							dateTimePicker_ngaysinh.CustomFormat = "dd/MM/yyyy";
						}
						else
						{
							dateTimePicker_ngaysinh.Checked = false;
							dateTimePicker_ngaysinh.CustomFormat = " ";
						}
					}
					else
					{
						dateTimePicker_ngaysinh.Checked = false;
						dateTimePicker_ngaysinh.CustomFormat = " ";
					}


					textBox_quequan.Text = selectedRow.Cells["DiaChi"]?.Value?.ToString() ?? "";
					textBox_quoctich.Text = selectedRow.Cells["QuocTich"]?.Value?.ToString() ?? "";
					textBox_email.Text = selectedRow.Cells["Email"]?.Value?.ToString() ?? "";
					textBox_trangthai.Text = selectedRow.Cells["TrangThai"]?.Value?.ToString() ?? "";
					textBox_tieusu.Text = selectedRow.Cells["TieuSu"]?.Value?.ToString() ?? "";

					// Sử dụng hàm SetComboBoxSelection đã sửa
					SetComboBoxSelection(comboBox_hocham, selectedRow.Cells["HocHam"]?.Value);
					SetComboBoxSelection(comboBox_hocvi, selectedRow.Cells["HocVi"]?.Value);

					LoadImageToPictureBox(currentImagePath);

					if (!string.IsNullOrEmpty(textBox_hovaten.Text))
					{
						LoadBooksBySelectedAuthor(textBox_hovaten.Text);
					}
					else
					{
						dataGridView_tacpham.DataSource = null;
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi hiển thị chi tiết tác giả: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error in DataGridView_tacgia_CellClick: {ex.StackTrace}");
					ClearSelectionAndFields();
				}
			}
			else
			{
				ClearSelectionAndFields();
			}
		}

		// Hàm xóa lựa chọn và các trường nhập liệu
		private void ClearSelectionAndFields()
		{
			selectedAuthorId = -1;
			selectedAuthorMaTacGia = null;
			currentImagePath = null;
			ClearInputFields();
			dataGridView_tacgia.ClearSelection();
			dataGridView_tacpham.DataSource = null;
		}

		// --- Sự kiện cho Nút Cập Nhật Ảnh ---
		private void button_capnhatanh_Click(object sender, EventArgs e)
		{
			if (selectedAuthorId <= 0 || string.IsNullOrEmpty(selectedAuthorMaTacGia))
			{
				MessageBox.Show("Vui lòng chọn một tác giả từ danh sách trước.", "Chưa chọn tác giả", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			try
			{
				using (OpenFileDialog openFileDialog = new OpenFileDialog())
				{
					openFileDialog.Filter = "Image Files|*.jpg;*.jpeg;*.png;*.gif;*.bmp";
					openFileDialog.Title = "Chọn ảnh đại diện cho tác giả";
					openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyPictures);

					if (openFileDialog.ShowDialog() == DialogResult.OK)
					{
						string selectedImagePath = openFileDialog.FileName;
						string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
						string targetDirectory = Path.Combine(documentsPath, "QLTV", "anhtacgia");
						Directory.CreateDirectory(targetDirectory);
						string extension = Path.GetExtension(selectedImagePath);
						string newFileName = $"{selectedAuthorMaTacGia}{extension}";
						string fullTargetPath = Path.Combine(targetDirectory, newFileName);

						File.Copy(selectedImagePath, fullTargetPath, true);

						string relativePathToSave = Path.Combine("anhtacgia", newFileName).Replace("\\", "/");

						if (modify.UpdateAuthorImage(selectedAuthorId, relativePathToSave)) // Giả sử hàm này tồn tại trong Modify.cs
						{
							MessageBox.Show("Cập nhật ảnh thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
							currentImagePath = relativePathToSave;
							LoadImageToPictureBox(currentImagePath);
						}
						else
						{
							MessageBox.Show("Cập nhật ảnh thất bại (Lỗi CSDL).", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
							try { if (File.Exists(fullTargetPath)) File.Delete(fullTargetPath); } catch { }
						}
					}
				}
			}
			catch (IOException ioEx)
			{
				MessageBox.Show($"Lỗi file khi cập nhật ảnh: {ioEx.Message}", "Lỗi File", MessageBoxButtons.OK, MessageBoxIcon.Error);
				Console.WriteLine($"File IO Error updating image: {ioEx.StackTrace}");
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi không mong muốn khi cập nhật ảnh: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
				Console.WriteLine($"Error updating image: {ex.StackTrace}");
			}
		}

		// --- Sự kiện cho Nút Cập Nhật Tiểu Sử ---
		private void button_capnhattieusu_Click(object sender, EventArgs e)
		{
			if (selectedAuthorId <= 0)
			{
				MessageBox.Show("Vui lòng chọn một tác giả từ danh sách trước.", "Chưa chọn tác giả", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			try
			{
				string newBiography = textBox_tieusu.Text;
				if (modify.UpdateAuthorBiography(selectedAuthorId, newBiography)) // Giả sử hàm này tồn tại trong Modify.cs
				{
					MessageBox.Show("Cập nhật tiểu sử thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				else
				{
					MessageBox.Show("Cập nhật tiểu sử thất bại.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi không mong muốn khi cập nhật tiểu sử: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
				Console.WriteLine($"Error updating biography: {ex.StackTrace}");
			}
		}

		// --- Sự kiện cho Nút Sửa (Cập nhật thông tin chung) ---
		private void button_sua_Click(object sender, EventArgs e)
		{
			if (selectedAuthorId <= 0)
			{
				MessageBox.Show("Vui lòng chọn một tác giả từ danh sách để cập nhật.", "Chưa chọn tác giả", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return;
			}

			try
			{
				TacGia updatedAuthor = new TacGia();
				updatedAuthor.ID = selectedAuthorId;
				updatedAuthor.MaTacGia = selectedAuthorMaTacGia;
				updatedAuthor.HoTen = textBox_hovaten.Text.Trim();

				if (dateTimePicker_ngaysinh.Checked && dateTimePicker_ngaysinh.Value != null)
				{
					updatedAuthor.NgaySinh = dateTimePicker_ngaysinh.Value.ToString("yyyy-MM-dd");
				}
				else
				{
					updatedAuthor.NgaySinh = null;
				}

				updatedAuthor.DiaChi = textBox_quequan.Text.Trim();
				updatedAuthor.QuocTich = textBox_quoctich.Text.Trim();
				updatedAuthor.Email = textBox_email.Text.Trim();
				updatedAuthor.TrangThai = textBox_trangthai.Text.Trim();
				updatedAuthor.HocHam = comboBox_hocham.SelectedItem?.ToString();
				updatedAuthor.HocVi = comboBox_hocvi.SelectedItem?.ToString();
				updatedAuthor.Anh = currentImagePath;
				updatedAuthor.TieuSu = textBox_tieusu.Text;

				if (modify.UpdateAuthorGeneralInfo(updatedAuthor)) // Giả sử hàm này tồn tại trong Modify.cs
				{
					MessageBox.Show("Cập nhật thông tin tác giả thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					int currentRowIndex = -1;
					if (dataGridView_tacgia.CurrentRow != null)
					{
						currentRowIndex = dataGridView_tacgia.CurrentRow.Index;
					}
					LoadAuthorDataGridView();
					if (currentRowIndex >= 0 && currentRowIndex < dataGridView_tacgia.Rows.Count)
					{
						dataGridView_tacgia.Rows[currentRowIndex].Selected = true;
						dataGridView_tacgia.CurrentCell = dataGridView_tacgia.Rows[currentRowIndex].Cells[1];
					}
				}
				else
				{
					MessageBox.Show("Cập nhật thông tin tác giả thất bại.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi không mong muốn khi cập nhật thông tin: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
				Console.WriteLine($"Error updating author info: {ex.StackTrace}");
			}
		}

		// --- Sự kiện cho Nút Lọc Tác Phẩm ---
		private void button_loc_Click(object sender, EventArgs e)
		{
			string authorNameToFilter = null;
			if (dataGridView_tacgia.SelectedRows.Count > 0)
			{
				authorNameToFilter = dataGridView_tacgia.SelectedRows[0].Cells["HoTen"]?.Value?.ToString();
			}
			else if (selectedAuthorId > 0 && !string.IsNullOrEmpty(textBox_hovaten.Text))
			{
				authorNameToFilter = textBox_hovaten.Text;
			}

			if (!string.IsNullOrEmpty(authorNameToFilter))
			{
				LoadBooksBySelectedAuthor(authorNameToFilter);
			}
			else
			{
				MessageBox.Show("Vui lòng chọn một tác giả trong danh sách hoặc đảm bảo tên tác giả đang hiển thị.", "Chưa chọn tác giả", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				LoadBookDataGridView();
			}
		}

		// --- Sự kiện cho Nút Làm Mới ---
		private void button_lammoi_Click(object sender, EventArgs e)
		{
			ClearSelectionAndFields();
			LoadAuthorDataGridView();
			LoadBookDataGridView();
		}


		// *** HÀM HELPER SETCOMBOBOXSELECTION ĐÃ SỬA LỖI ***
		/// <summary>
		/// Helper method để đặt giá trị cho ComboBox, kiểm tra null và tồn tại trong items.
		/// Đã sửa lỗi mơ hồ tham chiếu bằng cách chỉ định rõ namespace System.Windows.Forms.
		/// </summary>
		// *** SỬA LỖI: Chỉ định rõ namespace để giải quyết lỗi mơ hồ ***
		private void SetComboBoxSelection(System.Windows.Forms.ComboBox comboBox, object value)
		{
			comboBox.SelectedIndex = -1; // Reset trước
			if (value != null && value != DBNull.Value)
			{
				string valueString = value.ToString();
				int indexToSelect = -1;

				// Lặp qua các item trong ComboBox để tìm index khớp chính xác
				for (int i = 0; i < comboBox.Items.Count; i++)
				{
					// So sánh chuỗi một cách chính xác (phân biệt hoa thường)
					// Cần kiểm tra comboBox.Items[i] không null trước khi gọi ToString()
					if (comboBox.Items[i] != null && comboBox.Items[i].ToString().Equals(valueString))
					{
						indexToSelect = i;
						break; // Tìm thấy, thoát vòng lặp
					}
				}

				// Nếu tìm thấy index hợp lệ thì chọn nó
				if (indexToSelect >= 0)
				{
					comboBox.SelectedIndex = indexToSelect;
				}
				// Không cần else vì đã reset SelectedIndex = -1 ở đầu
			}
			// Nếu value là null hoặc DBNull, SelectedIndex đã được set là -1
		}

		// *** HÀM HELPER LOADIMAGE ***
		/// <summary>
		/// Helper method để tải ảnh từ đường dẫn vào PictureBox, xử lý lỗi.
		/// Đường dẫn ảnh được giả định là tương đối trong thư mục Documents/QLTV/anhtacgia
		/// </summary>
		private void LoadImageToPictureBox(string relativeImagePath)
		{
			pictureBox1.Image = null; // Xóa ảnh cũ
			if (string.IsNullOrEmpty(relativeImagePath))
			{
				return; // Không có đường dẫn, không làm gì cả
			}

			try
			{
				// Xây dựng đường dẫn tuyệt đối
				string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
				// Sử dụng Path.Combine để nối đường dẫn an toàn
				string fullPath = Path.Combine(documentsPath, "QLTV", relativeImagePath.Replace("/", "\\")); // Đảm bảo dùng đúng dấu \ cho Windows

				if (File.Exists(fullPath))
				{
					// Sử dụng MemoryStream để tránh khóa file ảnh
					byte[] imageBytes = File.ReadAllBytes(fullPath);
					using (var ms = new MemoryStream(imageBytes))
					{
						pictureBox1.Image = Image.FromStream(ms);
					}
					pictureBox1.SizeMode = PictureBoxSizeMode.Zoom; // Hoặc StretchImage, tùy bạn muốn
				}
				else
				{
					Console.WriteLine($"Không tìm thấy file ảnh tại đường dẫn đã giải quyết: {fullPath}");
					// Có thể hiển thị ảnh mặc định ở đây nếu muốn
					// pictureBox1.Image = Properties.Resources.DefaultAuthorImage;
				}
			}
			catch (FileNotFoundException)
			{
				Console.WriteLine($"Không tìm thấy file ảnh: {relativeImagePath}");
			}
			catch (OutOfMemoryException)
			{
				Console.WriteLine($"Tệp ảnh không hợp lệ hoặc quá lớn: {relativeImagePath}");
				MessageBox.Show("Tệp ảnh không hợp lệ hoặc quá lớn.", "Lỗi Ảnh", MessageBoxButtons.OK, MessageBoxIcon.Warning);
			}
			catch (Exception ex) // Bắt các lỗi khác (ví dụ: không đủ quyền đọc file)
			{
				Console.WriteLine($"Lỗi khi tải ảnh '{relativeImagePath}': {ex.Message}");
				// Không nên hiển thị MessageBox cho mọi lỗi nhỏ ở đây, chỉ log là đủ
			}
		}

		// *** HÀM HELPER CLEARINPUTFIELDS ***
		/// <summary>
		/// Xóa nội dung trên các control nhập liệu bên phải.
		/// </summary>
		private void ClearInputFields()
		{
			textBox_hovaten.Clear();
			dateTimePicker_ngaysinh.Checked = false;
			dateTimePicker_ngaysinh.CustomFormat = " "; // Hiển thị trống
			textBox_quequan.Clear();
			textBox_quoctich.Clear();
			textBox_email.Clear();
			textBox_trangthai.Clear();
			textBox_tieusu.Clear();
			comboBox_hocham.SelectedIndex = -1;
			comboBox_hocvi.SelectedIndex = -1;
			pictureBox1.Image = null;
		}

		// *** HÀM HELPER LOADBOOKSBYSELECTEDAUTHOR ***
		/// <summary>
		/// Tải danh sách sách của tác giả được chỉ định vào dataGridView_tacpham.
		/// </summary>
		/// <param name="authorName">Tên tác giả cần lọc.</param>
		private void LoadBooksBySelectedAuthor(string authorName)
		{
			try
			{
				DataTable booksData = modify.GetBooksByAuthorName(authorName); // Giả sử hàm này tồn tại trong Modify.cs
				dataGridView_tacpham.DataSource = null; // Xóa dữ liệu cũ
				dataGridView_tacpham.DataSource = booksData;

				// Cấu hình lại cột nếu cần
				if (dataGridView_tacpham.Columns.Count > 0)
				{
					try { dataGridView_tacpham.Columns["ID"].Visible = false; } catch { }
					try { dataGridView_tacpham.Columns["Masach"].HeaderText = "Mã sách"; } catch { }
					try { dataGridView_tacpham.Columns["TenSach"].HeaderText = "Tên sách"; } catch { }
					try { dataGridView_tacpham.Columns["TheLoai"].HeaderText = "Thể loại"; } catch { }
					try { dataGridView_tacpham.Columns["TacGia"].HeaderText = "Tác giả"; } catch { }
					try { dataGridView_tacpham.Columns["SoLuong"].HeaderText = "Số lượng"; } catch { }
					try { dataGridView_tacpham.Columns["NhaXuatBan"].HeaderText = "Nhà xuất bản"; } catch { }
					try { dataGridView_tacpham.Columns["NamXuatBan"].HeaderText = "Năm xuất bản"; } catch { }

					dataGridView_tacpham.Columns["TenSach"].AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
					dataGridView_tacpham.Columns["TacGia"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;

					dataGridView_tacpham.ReadOnly = true;
					dataGridView_tacpham.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
					dataGridView_tacpham.AllowUserToAddRows = false;
				}

				// Không cần thông báo nếu không tìm thấy sách, để grid trống là đủ
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi lọc tác phẩm của tác giả '{authorName}': {ex.Message}", "Lỗi Lọc Dữ Liệu", MessageBoxButtons.OK, MessageBoxIcon.Error);
				Console.WriteLine($"Error filtering books for author '{authorName}': {ex.StackTrace}");
				dataGridView_tacpham.DataSource = null;
			}
		}

		private void button_lammoi_Click_1(object sender, EventArgs e)
		{
			LoadBookDataGridView();
			LoadAuthorDataGridView();
			ClearInputFields();
		}
	} // Kết thúc lớp tacgia
} // Kết thúc namespace