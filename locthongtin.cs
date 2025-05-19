using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace ungdungthuviencaocap
{
	public partial class locthongtin : DevExpress.XtraEditors.XtraForm
	{
		private Modify modify;
		private timkiemsach formCha;

		public locthongtin()
		{
			try
			{
				InitializeComponent();

				// Khởi tạo đối tượng modify
				modify = new Modify();

				// Tìm form cha timkiemsach
				foreach (Form form in Application.OpenForms)
				{
					if (form is timkiemsach)
					{
						formCha = (timkiemsach)form;
						break;
					}
				}

				// Đăng ký các sự kiện
				this.Load += new EventHandler(locthongtin_Load);
				this.button_locthongtin.Click += new EventHandler(button_locthongtin_Click);
				this.button_loadlaiform.Click += new EventHandler(button_loadlaiform_Click);
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khởi tạo form: {ex.Message}",
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void locthongtin_Load(object sender, EventArgs e)
		{
			try
			{
				// Xóa và khởi tạo tất cả combobox với tùy chọn trống
				KhoiTaoComboBox();

				// Điền dữ liệu vào combobox từ cơ sở dữ liệu
				DienDuLieuComboBox();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tải dữ liệu lọc: {ex.Message}",
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void KhoiTaoComboBox()
		{
			try
			{
				// Xóa tất cả các mục trong combobox
				comboBox_tensach.Items.Clear();
				comboBox_theloai.Items.Clear();
				comboBox_tacgia.Items.Clear();
				comboBox_soluong.Items.Clear();
				comboBox_tunam.Items.Clear();
				comboBox_dennam.Items.Clear();
				comboBox_nhaxuatban.Items.Clear();

				// Thêm mục trống vào mỗi combobox
				string mucTrong = "-- Tất cả --";
				comboBox_tensach.Items.Add(mucTrong);
				comboBox_theloai.Items.Add(mucTrong);
				comboBox_tacgia.Items.Add(mucTrong);
				comboBox_soluong.Items.Add(mucTrong);
				comboBox_tunam.Items.Add(mucTrong);
				comboBox_dennam.Items.Add(mucTrong);
				comboBox_nhaxuatban.Items.Add(mucTrong);

				// Chọn mục trống làm mặc định
				comboBox_tensach.SelectedIndex = 0;
				comboBox_theloai.SelectedIndex = 0;
				comboBox_tacgia.SelectedIndex = 0;
				comboBox_soluong.SelectedIndex = 0;
				comboBox_tunam.SelectedIndex = 0;
				comboBox_dennam.SelectedIndex = 0;
				comboBox_nhaxuatban.SelectedIndex = 0;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi khởi tạo ComboBox: {ex.Message}",
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void DienDuLieuComboBox()
		{
			try
			{
				// Lấy và thêm các giá trị duy nhất từ mỗi cột
				ThemGiaTriDuyNhat(comboBox_tensach, "TenSach");
				ThemGiaTriDuyNhat(comboBox_theloai, "TheLoai");
				ThemGiaTriDuyNhat(comboBox_tacgia, "TacGia");
				ThemGiaTriDuyNhat(comboBox_soluong, "SoLuong");
				ThemGiaTriDuyNhat(comboBox_nhaxuatban, "NhaXuatBan");

				// Thêm các năm vào cả hai combobox năm
				List<string> danhSachNam = modify.GetDistinctValues("NamXuatBan");
				foreach (string nam in danhSachNam)
				{
					comboBox_tunam.Items.Add(nam);
					comboBox_dennam.Items.Add(nam);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi điền dữ liệu vào ComboBox: {ex.Message}",
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void ThemGiaTriDuyNhat(System.Windows.Forms.ComboBox comboBox, string tenCot)
		{
			try
			{
				List<string> danhSachGiaTri = modify.GetDistinctValues(tenCot);
				foreach (string giaTri in danhSachGiaTri)
				{
					comboBox.Items.Add(giaTri);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi thêm giá trị vào ComboBox {tenCot}: {ex.Message}",
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button_locthongtin_Click(object sender, EventArgs e)
		{
			try
			{
				if (formCha == null)
				{
					MessageBox.Show("Không thể tìm thấy form tìm kiếm sách.",
						"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				// Xây dựng câu truy vấn SQL với các điều kiện lọc
				string truyVan = XayDungCauTruyVanLoc();

				try
				{
					// Thực thi truy vấn và cập nhật lưới dữ liệu của form cha
					DataTable duLieuDaLoc = Connection.GetData(truyVan);
					CapNhatFormCha(duLieuDaLoc);

					this.DialogResult = DialogResult.OK;
				}
				catch (SQLiteException sqlEx)
				{
					MessageBox.Show($"Lỗi truy vấn cơ sở dữ liệu: {sqlEx.Message}",
						"Lỗi Database", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi áp dụng bộ lọc: {ex.Message}",
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private string XayDungCauTruyVanLoc()
		{
			try
			{
				StringBuilder dauRaSQL = new StringBuilder();
				List<string> dieuKien = new List<string>();

				// Thêm điều kiện cho mỗi bộ lọc được chọn
				if (comboBox_tensach.SelectedIndex > 0)
					dieuKien.Add($"TenSach = '{comboBox_tensach.Text.Replace("'", "''")}'");

				if (comboBox_theloai.SelectedIndex > 0)
					dieuKien.Add($"TheLoai = '{comboBox_theloai.Text.Replace("'", "''")}'");

				if (comboBox_tacgia.SelectedIndex > 0)
					dieuKien.Add($"TacGia = '{comboBox_tacgia.Text.Replace("'", "''")}'");

				if (comboBox_soluong.SelectedIndex > 0)
					dieuKien.Add($"SoLuong = '{comboBox_soluong.Text}'");

				if (comboBox_nhaxuatban.SelectedIndex > 0)
					dieuKien.Add($"NhaXuatBan = '{comboBox_nhaxuatban.Text.Replace("'", "''")}'");

				// Xử lý các bộ lọc về khoảng năm
				if (comboBox_tunam.SelectedIndex > 0 && comboBox_dennam.SelectedIndex > 0)
				{
					int tuNam = int.Parse(comboBox_tunam.Text);
					int denNam = int.Parse(comboBox_dennam.Text);
					dieuKien.Add($"NamXuatBan BETWEEN {tuNam} AND {denNam}");
				}
				else if (comboBox_tunam.SelectedIndex > 0)
				{
					dieuKien.Add($"NamXuatBan >= {comboBox_tunam.Text}");
				}
				else if (comboBox_dennam.SelectedIndex > 0)
				{
					dieuKien.Add($"NamXuatBan <= {comboBox_dennam.Text}");
				}

				// Xây dựng truy vấn cuối cùng
				string truyVanCoBan = "SELECT ID, Masach, TenSach, TheLoai, TacGia, SoLuong, " +
									 "NhaXuatBan, NamXuatBan, anh, pdf, tomtatnoidung FROM quanlysach";

				if (dieuKien.Count > 0)
					truyVanCoBan += " WHERE " + string.Join(" AND ", dieuKien);

				truyVanCoBan += " ORDER BY TenSach ASC";

				return truyVanCoBan;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi xây dựng câu truy vấn: {ex.Message}",
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return "SELECT ID, Masach, TenSach, TheLoai, TacGia, SoLuong, NhaXuatBan, NamXuatBan, anh, pdf, tomtatnoidung FROM quanlysach";
			}
		}

		private void CapNhatFormCha(DataTable duLieuDaLoc)
		{
			try
			{
				// Chúng ta cần thêm phương thức này vào lớp timkiemsach
				if (formCha != null && duLieuDaLoc != null)
				{
					formCha.CapNhatDataGridVoiKetQuaLoc(duLieuDaLoc);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi cập nhật form cha: {ex.Message}",
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button_loadlaiform_Click(object sender, EventArgs e)
		{
			try
			{
				if (formCha == null)
				{
					MessageBox.Show("Không thể tìm thấy form tìm kiếm sách.",
						"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				// Đặt lại dữ liệu trong form cha
				formCha.TaiLaiDuLieuGoc();

				this.DialogResult = DialogResult.OK;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tải lại dữ liệu: {ex.Message}",
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
