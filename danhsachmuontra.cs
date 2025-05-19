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
	public partial class danhsachmuontra : DevExpress.XtraEditors.XtraForm
	{
		public danhsachmuontra()
		{
			InitializeComponent();
			modify = new Modify();
		}

		Modify modify;

		private void danhsachmuontra_Load(object sender, EventArgs e)
		{
			try
			{
				dataGridView1.DataSource = modify.getAllphieu();
				dataGridView1.Columns["ID"].Visible = false;
				dataGridView1.Columns["IsActive"].Visible = false;
				dataGridView1.Columns["TenSach"].HeaderText = "Tên sách";
				dataGridView1.Columns["MaSinhVien"].HeaderText = "Mã sinh viên";
				dataGridView1.Columns["NgayMuon"].HeaderText = "Ngày mượn";
				dataGridView1.Columns["NgayTra"].HeaderText = "Ngày trả";
				dataGridView1.Columns["HoVaTen"].HeaderText = "Họ và tên";
				dataGridView1.Columns["LoaiPhieu"].HeaderText = "Loại phiếu";
				dataGridView1.Columns["SoLuong"].HeaderText = "Số lượng";
				dataGridView1.Columns["MaSach"].HeaderText = "Mã sách";
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}