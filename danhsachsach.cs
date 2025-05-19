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
	public partial class danhsachsach: DevExpress.XtraEditors.XtraForm
	{
        public danhsachsach()
		{
            InitializeComponent();
		}
		Modify modify;
		private void danhsachsach_Load(object sender, EventArgs e)
		{
			modify = new Modify();
			try
			{
				dataGridView1.DataSource = modify.getAllbook();
				dataGridView1.Columns["ID"].Visible = false;
				dataGridView1.Columns["anh"].Visible = false;
				dataGridView1.Columns["pdf"].Visible = false;
				dataGridView1.Columns["tomtatnoidung"].Visible = false;
				dataGridView1.Columns["Masach"].HeaderText = "Mã sách";
				dataGridView1.Columns["TenSach"].HeaderText = "Tên sách";
				dataGridView1.Columns["TheLoai"].HeaderText = "Thể loại";
				dataGridView1.Columns["TacGia"].HeaderText = "Tác giả";
				dataGridView1.Columns["SoLuong"].HeaderText = "Số lượng";
				dataGridView1.Columns["NhaXuatBan"].HeaderText = "Nhà xuất bản";
				dataGridView1.Columns["NamXuatBan"].HeaderText = "Năm xuất bản";
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}