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
	public partial class lichhenmuonsach : DevExpress.XtraEditors.XtraForm
	{
		public lichhenmuonsach()
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

		Modify modify;
		private void lichhenmuonsach_Load(object sender, EventArgs e)
		{
			modify = new Modify();
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

		private void button_kiemtra_Click(object sender, EventArgs e)
		{
			try
			{
				string timsv = this.textBox1.Text;
				using (SQLiteConnection sqliteConnection = Connection.GetSQLiteConnection())
				{
					sqliteConnection.Open();
					dataGridView1.DataSource = modify.timsvv(timsv, sqliteConnection);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tìm kiếm sinh viên: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button3_Click(object sender, EventArgs e)
		{
			try
			{
				if (dataGridView1.SelectedRows.Count == 0)
				{
					MessageBox.Show("Bạn chưa chọn để xác nhận", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				int ID = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells[0].Value);
				DialogResult result = MessageBox.Show("Bạn đã chắc chắn họ đến đúng hẹn rồi chứ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
				if (result == DialogResult.Yes)
				{
					if (modify.dunghen(ID))
					{
						MessageBox.Show("Xác nhận thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
						dataGridView1.DataSource = modify.getAllhen();
					}
					else
					{
						MessageBox.Show("Xác nhận thất bại", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi xác nhận đúng hẹn: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button1_Click(object sender, EventArgs e)
		{
			try
			{
				if (dataGridView1.SelectedRows.Count == 0)
				{
					MessageBox.Show("Bạn chưa chọn để xác nhận", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return;
				}

				int ID = Convert.ToInt32(dataGridView1.SelectedRows[0].Cells[0].Value);
				DialogResult result = MessageBox.Show("Bạn đã chắc chắn họ quá hẹn rồi chứ?", "Thông báo", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
				if (result == DialogResult.Yes)
				{
					if (modify.dunghen(ID))
					{
						MessageBox.Show("Xác nhận thành công", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
						dataGridView1.DataSource = modify.getAllhen();
					}
					else
					{
						MessageBox.Show("Xác nhận thất bại", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi xác nhận quá hẹn: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void textBox1_KeyDown(object sender, KeyEventArgs e)
		{
			try
			{
				if (e.KeyCode == Keys.Enter)
				{
					button_kiemtra_Click(sender, e);
					e.Handled = true;
					e.SuppressKeyPress = true;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi xử lý phím Enter: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
