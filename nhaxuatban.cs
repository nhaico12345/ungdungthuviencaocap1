using DevExpress.XtraEditors;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ungdungthuviencaocap
{
	public partial class nhaxuatban : DevExpress.XtraEditors.XtraForm
	{
		Modify modify;
		private int currentNXB_ID = 0;

		public nhaxuatban()
		{
			InitializeComponent();
			this.dataGridView_nhaxb.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_nhaxb_CellClick);
			this.button_capnhat.Click += new System.EventHandler(this.button_capnhat_Click);
			this.button_lammoi.Click += new System.EventHandler(this.button_lammoi_Click);
			this.button_loc.Click += new System.EventHandler(this.button_loc_Click); // Đăng ký sự kiện cho nút lọc
		}

		private void nhaxuatban_Load(object sender, EventArgs e)
		{
			try
			{
				modify = new Modify();
				LoadNhaXuatBanDataGridView();
				LoadBookDataGridView();

				dateTimePicker_ngaythanhlap.CustomFormat = "dd/MM/yyyy";
				dateTimePicker_ngaythanhlap.Format = DateTimePickerFormat.Custom;
				dateTimePicker_ngaythanhlap.Value = DateTime.Now.Date;
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi tải form Nhà xuất bản: " + ex.Message, "Lỗi Khởi Tạo", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void LoadNhaXuatBanDataGridView()
		{
			try
			{
				dataGridView_nhaxb.DataSource = null;
				DataTable dtNXB = modify.GetAllNhaXuatBan();
				dataGridView_nhaxb.DataSource = dtNXB;

				if (dataGridView_nhaxb.Columns.Count > 0)
				{
					if (dataGridView_nhaxb.Columns.Contains("ID"))
						dataGridView_nhaxb.Columns["ID"].Visible = false;
					if (dataGridView_nhaxb.Columns.Contains("MaNhaXuatBan"))
						dataGridView_nhaxb.Columns["MaNhaXuatBan"].HeaderText = "Mã NXB";
					if (dataGridView_nhaxb.Columns.Contains("TenNhaXuatBan"))
						dataGridView_nhaxb.Columns["TenNhaXuatBan"].HeaderText = "Tên Nhà Xuất Bản";
					if (dataGridView_nhaxb.Columns.Contains("DiaChi"))
						dataGridView_nhaxb.Columns["DiaChi"].HeaderText = "Địa Chỉ";
					if (dataGridView_nhaxb.Columns.Contains("SoHieuXuatBan"))
						dataGridView_nhaxb.Columns["SoHieuXuatBan"].HeaderText = "Số Hiệu XB";
					if (dataGridView_nhaxb.Columns.Contains("Email"))
						dataGridView_nhaxb.Columns["Email"].HeaderText = "Email";
					if (dataGridView_nhaxb.Columns.Contains("Website"))
						dataGridView_nhaxb.Columns["Website"].HeaderText = "Website";
					if (dataGridView_nhaxb.Columns.Contains("NgayThanhLap"))
					{
						dataGridView_nhaxb.Columns["NgayThanhLap"].HeaderText = "Ngày Thành Lập";
						// Định dạng cột ngày tháng nếu cần (ví dụ: dd/MM/yyyy)
						dataGridView_nhaxb.Columns["NgayThanhLap"].DefaultCellStyle.Format = "dd/MM/yyyy";
					}
					if (dataGridView_nhaxb.Columns.Contains("TrangThai"))
						dataGridView_nhaxb.Columns["TrangThai"].HeaderText = "Trạng Thái";

					dataGridView_nhaxb.ReadOnly = true;
					dataGridView_nhaxb.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
					dataGridView_nhaxb.AllowUserToAddRows = false;
					dataGridView_nhaxb.MultiSelect = false;
					dataGridView_nhaxb.RowHeadersVisible = false;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi tải dữ liệu Nhà Xuất Bản vào GridView: " + ex.Message, "Lỗi Tải Dữ Liệu", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}


		private void LoadBookDataGridView(DataTable bookData = null)
		{
			try
			{
				dataGridView_tacpham.DataSource = null;
				if (bookData == null)
				{
					dataGridView_tacpham.DataSource = modify.getAllbooks();
				}
				else
				{
					dataGridView_tacpham.DataSource = bookData;
				}

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
					if (dataGridView_tacpham.Columns.Contains("anh")) dataGridView_tacpham.Columns["anh"].Visible = false;
					if (dataGridView_tacpham.Columns.Contains("pdf")) dataGridView_tacpham.Columns["pdf"].Visible = false;
					if (dataGridView_tacpham.Columns.Contains("tomtatnoidung")) dataGridView_tacpham.Columns["tomtatnoidung"].Visible = false;

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

		private void dataGridView_nhaxb_CellClick(object sender, DataGridViewCellEventArgs e)
		{
			if (e.RowIndex >= 0 && e.RowIndex < dataGridView_nhaxb.Rows.Count && !dataGridView_nhaxb.Rows[e.RowIndex].IsNewRow)
			{
				DataGridViewRow row = dataGridView_nhaxb.Rows[e.RowIndex];
				if (row.Cells["ID"].Value != null && int.TryParse(row.Cells["ID"].Value.ToString(), out int id))
				{
					currentNXB_ID = id;
				}
				else
				{
					currentNXB_ID = 0;
				}

				textBox_tennhaxuatban.Text = row.Cells["TenNhaXuatBan"]?.Value?.ToString() ?? "";
				textBox_diachi.Text = row.Cells["DiaChi"]?.Value?.ToString() ?? "";
				textBox_sohieuxuatban.Text = row.Cells["SoHieuXuatBan"]?.Value?.ToString() ?? "";
				textBox_email.Text = row.Cells["Email"]?.Value?.ToString() ?? "";
				textBox_websiste.Text = row.Cells["Website"]?.Value?.ToString() ?? "";
				textBox_trangthai.Text = row.Cells["TrangThai"]?.Value?.ToString() ?? "";

				string ngayThanhLapStr = row.Cells["NgayThanhLap"]?.Value?.ToString();
				if (!string.IsNullOrEmpty(ngayThanhLapStr))
				{
					if (DateTime.TryParseExact(ngayThanhLapStr, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime ngayTL) ||
						DateTime.TryParseExact(ngayThanhLapStr, "dd/MM/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out ngayTL)) // Thêm dd/MM/yyyy
					{
						dateTimePicker_ngaythanhlap.Value = ngayTL;
					}
					else
					{
						dateTimePicker_ngaythanhlap.Value = DateTime.Now.Date;
						Console.WriteLine($"Không thể parse ngày: {ngayThanhLapStr}");
					}
				}
				else
				{
					dateTimePicker_ngaythanhlap.Value = DateTime.Now.Date;
				}
			}
			else
			{
				currentNXB_ID = 0;
			}
		}

		private void button_capnhat_Click(object sender, EventArgs e)
		{
			string tenNXB_form = textBox_tennhaxuatban.Text.Trim();
			string diaChi_form = textBox_diachi.Text.Trim();
			string soHieu_form = textBox_sohieuxuatban.Text.Trim();
			string email_form = textBox_email.Text.Trim();
			string website_form = textBox_websiste.Text.Trim();
			string trangThai_form = textBox_trangthai.Text.Trim();
			string ngayTL_form = dateTimePicker_ngaythanhlap.Value.ToString("yyyy-MM-dd");

			if (string.IsNullOrWhiteSpace(tenNXB_form))
			{
				MessageBox.Show("Tên nhà xuất bản không được để trống.", "Thiếu Thông Tin", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				textBox_tennhaxuatban.Focus();
				return;
			}
			// Thêm các kiểm tra khác cho các trường bắt buộc nếu có

			try
			{
				if (currentNXB_ID == 0) // Trường hợp thêm mới (không có dòng nào được chọn)
				{
					NhaXuatBan existingNXB = modify.GetNhaXuatBanByName(tenNXB_form);
					if (existingNXB != null)
					{
						// Tên NXB đã tồn tại
						DialogResult dr = MessageBox.Show($"Nhà xuất bản '{tenNXB_form}' đã tồn tại trong hệ thống. Bạn có muốn cập nhật thông tin cho nhà xuất bản này không?", "Xác Nhận Ghi Đè", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
						if (dr == DialogResult.Yes)
						{
							// Cập nhật NXB hiện có
							existingNXB.DiaChi = diaChi_form;
							existingNXB.SoHieuXuatBan = soHieu_form;
							existingNXB.Email = email_form;
							existingNXB.Website = website_form;
							existingNXB.NgayThanhLap = ngayTL_form;
							existingNXB.TrangThai = trangThai_form;
							// MaNhaXuatBan không thay đổi khi cập nhật theo tên

							if (modify.UpdateNhaXuatBan(existingNXB))
							{
								MessageBox.Show("Cập nhật thông tin nhà xuất bản thành công!", "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
								LoadNhaXuatBanDataGridView();
								ClearForm();
							}
						}
						// Nếu No, không làm gì cả
					}
					else
					{
						// Tên NXB chưa tồn tại, thêm mới
						DialogResult dr = MessageBox.Show($"Bạn có chắc chắn muốn thêm nhà xuất bản mới '{tenNXB_form}' không?", "Xác Nhận Thêm Mới", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
						if (dr == DialogResult.Yes)
						{
							NhaXuatBan newNXB = new NhaXuatBan
							{
								// MaNhaXuatBan sẽ được tạo tự động trong InsertNhaXuatBan
								TenNhaXuatBan = tenNXB_form,
								DiaChi = diaChi_form,
								SoHieuXuatBan = soHieu_form,
								Email = email_form,
								Website = website_form,
								NgayThanhLap = ngayTL_form,
								TrangThai = trangThai_form
							};
							if (modify.InsertNhaXuatBan(newNXB, true)) // true để tự tạo mã
							{
								MessageBox.Show("Thêm nhà xuất bản mới thành công!", "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
								LoadNhaXuatBanDataGridView();
								ClearForm();
							}
						}
						// Nếu No, không làm gì cả
					}
				}
				else // Trường hợp cập nhật (đã chọn một dòng)
				{
					string maNXB_selected = dataGridView_nhaxb.CurrentRow.Cells["MaNhaXuatBan"].Value.ToString();

					// Kiểm tra xem tên NXB mới có trùng với NXB khác không
					NhaXuatBan nxbWithSameName = modify.GetNhaXuatBanByName(tenNXB_form);
					if (nxbWithSameName != null && nxbWithSameName.ID != currentNXB_ID)
					{
						MessageBox.Show($"Tên nhà xuất bản '{tenNXB_form}' đã được sử dụng bởi một nhà xuất bản khác (Mã: {nxbWithSameName.MaNhaXuatBan}). Vui lòng chọn tên khác.", "Lỗi Trùng Lặp", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}

					NhaXuatBan nxbToUpdate = new NhaXuatBan
					{
						ID = currentNXB_ID,
						MaNhaXuatBan = maNXB_selected, // Giữ nguyên mã NXB khi cập nhật
						TenNhaXuatBan = tenNXB_form,
						DiaChi = diaChi_form,
						SoHieuXuatBan = soHieu_form,
						Email = email_form,
						Website = website_form,
						NgayThanhLap = ngayTL_form,
						TrangThai = trangThai_form
					};

					if (modify.UpdateNhaXuatBan(nxbToUpdate))
					{
						MessageBox.Show("Cập nhật thông tin nhà xuất bản thành công!", "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
						LoadNhaXuatBanDataGridView();
						ClearForm();
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi không mong muốn: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
				Console.WriteLine($"Stack Trace lỗi: {ex.StackTrace}");
			}
		}

		private void ClearForm()
		{
			currentNXB_ID = 0;
			textBox_tennhaxuatban.Clear();
			textBox_diachi.Clear();
			textBox_sohieuxuatban.Clear();
			textBox_email.Clear();
			textBox_websiste.Clear();
			textBox_trangthai.Clear();
			dateTimePicker_ngaythanhlap.Value = DateTime.Now.Date;
			dataGridView_nhaxb.ClearSelection();
			textBox_tennhaxuatban.Focus();
		}

		private void button_lammoi_Click(object sender, EventArgs e)
		{
			ClearForm();
			LoadNhaXuatBanDataGridView();
			LoadBookDataGridView();
		}

		private void button_loc_Click(object sender, EventArgs e)
		{
			if (dataGridView_nhaxb.CurrentRow != null && dataGridView_nhaxb.CurrentRow.Index >= 0)
			{
				string tenNhaXuatBanCanLoc = dataGridView_nhaxb.CurrentRow.Cells["TenNhaXuatBan"].Value?.ToString();
				if (!string.IsNullOrEmpty(tenNhaXuatBanCanLoc))
				{
					DataTable filteredBooks = modify.GetBooksByPublisherName(tenNhaXuatBanCanLoc);
					LoadBookDataGridView(filteredBooks);
					if (filteredBooks.Rows.Count == 0)
					{
						MessageBox.Show($"Không tìm thấy sách nào của nhà xuất bản '{tenNhaXuatBanCanLoc}'.", "Thông Báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				}
				else
				{
					MessageBox.Show("Không thể lấy tên nhà xuất bản để lọc.", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			else
			{
				MessageBox.Show("Vui lòng chọn một nhà xuất bản từ danh sách để lọc sách.", "Chưa Chọn Nhà Xuất Bản", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				LoadBookDataGridView(); // Hiển thị lại tất cả sách nếu không có NXB nào được chọn
			}
		}
	}
}
