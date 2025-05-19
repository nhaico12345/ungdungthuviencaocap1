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
	public partial class thongkethongtinsach: DevExpress.XtraEditors.XtraForm
	{
		private Modify modify = new Modify();
		public thongkethongtinsach()
		{
            InitializeComponent();
		}

		private void button_timtheloai_Click(object sender, EventArgs e)
		{
			try
			{
				string theLoai = comboBox_theloaisach.Text;

				if (string.IsNullOrEmpty(theLoai))
				{
					MessageBox.Show("Vui lòng chọn hoặc nhập thể loại cần tìm!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				DataTable result = modify.SearchBooksByColumn("TheLoai", theLoai);

				if (result.Rows.Count == 0)
				{
					MessageBox.Show("Thể loại sách này chưa có hoặc sách về thể loại này chưa có, vui lòng đợi thêm!",
									"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				dataGridView1.DataSource = result;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tìm kiếm thể loại: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button_timtacgia_Click(object sender, EventArgs e)
		{
			try
			{
				string tacGia = comboBox_tacgia.Text;

				if (string.IsNullOrEmpty(tacGia))
				{
					MessageBox.Show("Vui lòng chọn hoặc nhập tác giả cần tìm!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				DataTable result = modify.SearchBooksByColumn("TacGia", tacGia);

				if (result.Rows.Count == 0)
				{
					MessageBox.Show("Tên tác giả này chưa có hoặc sách của Tác giả này chưa được thêm, vui lòng đợi thêm!",
									"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				dataGridView1.DataSource = result;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tìm kiếm tác giả: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button_timnhaxuatban_Click(object sender, EventArgs e)
		{
			try
			{
				string nhaXuatBan = comboBox_nhaxuatban.Text;

				if (string.IsNullOrEmpty(nhaXuatBan))
				{
					MessageBox.Show("Vui lòng chọn hoặc nhập nhà xuất bản cần tìm!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				DataTable result = modify.SearchBooksByColumn("NhaXuatBan", nhaXuatBan);

				if (result.Rows.Count == 0)
				{
					MessageBox.Show("Nhà xuất bản này chưa có hoặc sách của nhà xuất bản này chưa được thêm, vui lòng đợi thêm!",
									"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				dataGridView1.DataSource = result;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tìm kiếm nhà xuất bản: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button_namxuatban_Click(object sender, EventArgs e)
		{
			try
			{
				string namXuatBan = comboBox_namxuatban.Text;

				if (string.IsNullOrEmpty(namXuatBan))
				{
					MessageBox.Show("Vui lòng chọn hoặc nhập năm xuất bản cần tìm!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				// Kiểm tra nhập liệu năm để đảm bảo là số
				int year;
				if (!int.TryParse(namXuatBan, out year))
				{
					MessageBox.Show("Năm xuất bản phải là số!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				DataTable result = modify.SearchBooksByColumn("NamXuatBan", namXuatBan);

				if (result.Rows.Count == 0)
				{
					MessageBox.Show("Năm xuất bản này chưa có hoặc sách sản xuất năm này chưa được thêm, vui lòng đợi thêm!",
									"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				dataGridView1.DataSource = result;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tìm kiếm năm xuất bản: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button_soluong_Click(object sender, EventArgs e)
		{
			try
			{
				string soLuong = comboBox_soluong.Text;

				if (string.IsNullOrEmpty(soLuong))
				{
					MessageBox.Show("Vui lòng chọn hoặc nhập số lượng cần tìm!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				// Kiểm tra nhập liệu số lượng để đảm bảo là số
				int quantity;
				if (!int.TryParse(soLuong, out quantity))
				{
					MessageBox.Show("Số lượng phải là số!", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				DataTable result = modify.SearchBooksByColumn("SoLuong", soLuong);

				if (result.Rows.Count == 0)
				{
					MessageBox.Show("Số lượng sách này chưa có hoặc sách với số lượng này chưa được thêm, vui lòng đợi thêm!",
									"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return;
				}

				dataGridView1.DataSource = result;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tìm kiếm số lượng: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void thongkethongtinsach_Load(object sender, EventArgs e)
		{
			try
			{
				// Khởi đầu tải tất cả dữ liệu sách
				dataGridView1.DataSource = modify.getAllbooks();
				dataGridView1.Columns["anh"].Visible=false;
				dataGridView1.Columns["pdf"].Visible = false;
				dataGridView1.Columns["tomtatnoidung"].Visible = false;
				dataGridView1.Columns["ID"].Visible = false;
				dataGridView1.Columns["TenSach"].HeaderText = "Tên sách";
				dataGridView1.Columns["TheLoai"].HeaderText = "Thể loại";
				dataGridView1.Columns["TacGia"].HeaderText = "Tác giả";
				dataGridView1.Columns["NhaXuatBan"].HeaderText = "Nhà xuất bản";
				dataGridView1.Columns["NamXuatBan"].HeaderText = "Năm xuất bản";
				dataGridView1.Columns["SoLuong"].HeaderText = "Số lượng";

				// Tải các giá trị riêng biệt vào ComboBoxes
				LoadComboBoxes();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tải form: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void LoadComboBoxes()
		{
			try
			{
				// Tải thể loại
				List<string> genres = modify.GetDistinctValues("TheLoai");
				comboBox_theloaisach.Items.Clear();
				comboBox_theloaisach.Items.AddRange(genres.ToArray());

				// Tải tác giả
				List<string> authors = modify.GetDistinctValues("TacGia");
				comboBox_tacgia.Items.Clear();
				comboBox_tacgia.Items.AddRange(authors.ToArray());

				// Tải nhà xuất bản
				List<string> publishers = modify.GetDistinctValues("NhaXuatBan");
				comboBox_nhaxuatban.Items.Clear();
				comboBox_nhaxuatban.Items.AddRange(publishers.ToArray());

				// Tải năm xuất bản
				List<string> years = modify.GetDistinctValues("NamXuatBan");
				comboBox_namxuatban.Items.Clear();
				comboBox_namxuatban.Items.AddRange(years.ToArray());

				// Tải số lượng
				List<string> quantities = modify.GetDistinctValues("SoLuong");
				comboBox_soluong.Items.Clear();
				comboBox_soluong.Items.AddRange(quantities.ToArray());
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tải dữ liệu vào combobox: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button1_Click(object sender, EventArgs e)
		{
			// Khởi đầu tải tất cả dữ liệu sách
			dataGridView1.DataSource = modify.getAllbooks();
		}
	}
}