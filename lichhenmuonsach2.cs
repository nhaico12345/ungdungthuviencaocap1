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
	public partial class lichhenmuonsach2 : DevExpress.XtraEditors.XtraForm
	{
		public lichhenmuonsach2()
		{
			try
			{
				InitializeComponent();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khởi tạo form: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void LoadDataGridView()
		{
			try
			{
				dataGridView1.DataSource = modify.getAllhen();
				dataGridView1.Columns["ID"].Visible = false;
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button3_Click(object sender, EventArgs e)
		{
			try
			{
				string tensach = this.comboBox_tensach.Text;
				string masinhvien = this.textBox_masinhvien.Text;
				string hovaten = this.textBox_hovaten.Text;
				string loaihen = this.comboBox_muontrasach.Text;
				string giohen = this.textBox1.Text;
				DateTime lichhen = dateTimePicker_lichhen.Value;

				// Kiểm tra dữ liệu đầu vào
				if (string.IsNullOrEmpty(tensach))
				{
					MessageBox.Show("Vui lòng chọn tên sách", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				if (string.IsNullOrEmpty(masinhvien))
				{
					MessageBox.Show("Mã sinh viên không được để trống", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				if (string.IsNullOrEmpty(hovaten))
				{
					MessageBox.Show("Họ và tên không được để trống", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				if (string.IsNullOrEmpty(loaihen))
				{
					MessageBox.Show("Vui lòng chọn loại hẹn", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				if (string.IsNullOrEmpty(giohen))
				{
					MessageBox.Show("Giờ hẹn không được để trống", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				henmuonsach henmuonsach = new henmuonsach(0, masinhvien, hovaten, tensach, lichhen, giohen, loaihen);
				if (modify.datlichhen(henmuonsach))
				{
					MessageBox.Show("Thêm phiếu hẹn thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
					LoadDataGridView();
				}
				else
				{
					MessageBox.Show("Thêm phiếu hẹn thất bại", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi đặt lịch hẹn: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		Modify modify;
		private void lichhenmuonsach2_Load(object sender, EventArgs e)
		{
			try
			{
				modify = new Modify();
				LoadDataGridView();

				// Hiển thị thông tin người dùng đã đăng nhập
				if (UserSession.CurrentUser != null)
				{
					textBox_masinhvien.Text = UserSession.CurrentUser.MaSinhVien;
					textBox_hovaten.Text = UserSession.CurrentUser.HoVaTen;
				}
				else
				{
					MessageBox.Show("Không tìm thấy thông tin người dùng đăng nhập!",
						"Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}

				// Đảm bảo các textbox ở chế độ chỉ đọc
				textBox_masinhvien.ReadOnly = true;
				textBox_hovaten.ReadOnly = true;

				// Load danh sách sách
				DataTable dt = new DataTable();
				using (SQLiteConnection con = Connection.GetSQLiteConnection())
				{
					con.Open();
					SQLiteCommand command = new SQLiteCommand();
					command.CommandText = "select * from quanlysach";
					command.Connection = con;
					SQLiteDataReader readSQL = command.ExecuteReader();
					dt.Load(readSQL);
					comboBox_tensach.DataSource = dt;
					comboBox_tensach.DisplayMember = "Tensach";
					comboBox_tensach.ValueMember = "tensach";
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tải form lịch hẹn: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
