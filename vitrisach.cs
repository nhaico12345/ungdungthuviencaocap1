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
	public partial class vitrisach : DevExpress.XtraEditors.XtraForm
	{
		private Modify modify = new Modify();
		public vitrisach()
		{
			try
			{
				InitializeComponent();
				this.Load += vitrisach_Load;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khởi tạo form: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button_themvitri_Click(object sender, EventArgs e)
		{
			try
			{
				string selectedBook = comboBox_chonsachthem.SelectedValue?.ToString();
				string bookLocation = textBox_nhapvitrisach.Text.Trim();

				if (string.IsNullOrEmpty(selectedBook) || string.IsNullOrEmpty(bookLocation))
				{
					MessageBox.Show("Vui lòng chọn sách và nhập vị trí!", "Thông báo",
						MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
				{
					sqlConnection.Open();

					// Kiểm tra xem sách đã có vị trí chưa
					string checkQuery = "SELECT COUNT(*) FROM vitrisach WHERE TenSach = @TenSach";
					SQLiteCommand checkCommand = new SQLiteCommand(checkQuery, sqlConnection);
					checkCommand.Parameters.AddWithValue("@TenSach", selectedBook);

					int count = Convert.ToInt32(checkCommand.ExecuteScalar());
					if (count > 0)
					{
						MessageBox.Show("Sách này đã có vị trí! Vui lòng sử dụng chức năng sửa vị trí.",
							"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}

					// Chèn vị trí mới
					string insertQuery = "INSERT INTO vitrisach (TenSach, ViTri) VALUES (@TenSach, @ViTri)";
					SQLiteCommand insertCommand = new SQLiteCommand(insertQuery, sqlConnection);
					insertCommand.Parameters.AddWithValue("@TenSach", selectedBook);
					insertCommand.Parameters.AddWithValue("@ViTri", bookLocation);

					int result = insertCommand.ExecuteNonQuery();
					if (result > 0)
					{
						MessageBox.Show("Thêm vị trí sách thành công!", "Thông báo",
							MessageBoxButtons.OK, MessageBoxIcon.Information);

						// Làm mới dữ liệu
						LoadDataGridView();
						LoadBookLocations();
						textBox_nhapvitrisach.Clear();
					}
					else
					{
						MessageBox.Show("Thêm vị trí sách không thành công!", "Thông báo",
							MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button_suavitrisach_Click(object sender, EventArgs e)
		{
			try
			{
				if (dataGridView1.CurrentRow == null)
				{
					MessageBox.Show("Vui lòng chọn một vị trí sách để sửa!", "Thông báo",
						MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				string selectedBook = comboBox_chonsachthem.SelectedValue?.ToString();
				string bookLocation = textBox_nhapvitrisach.Text.Trim();

				if (string.IsNullOrEmpty(selectedBook) || string.IsNullOrEmpty(bookLocation))
				{
					MessageBox.Show("Vui lòng chọn sách và nhập vị trí!", "Thông báo",
						MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
				{
					sqlConnection.Open();
					string updateQuery = "UPDATE vitrisach SET ViTri = @ViTri WHERE TenSach = @TenSach";
					SQLiteCommand updateCommand = new SQLiteCommand(updateQuery, sqlConnection);
					updateCommand.Parameters.AddWithValue("@TenSach", selectedBook);
					updateCommand.Parameters.AddWithValue("@ViTri", bookLocation);

					int result = updateCommand.ExecuteNonQuery();
					if (result > 0)
					{
						MessageBox.Show("Cập nhật vị trí sách thành công!", "Thông báo",
							MessageBoxButtons.OK, MessageBoxIcon.Information);

						LoadDataGridView();
						textBox_nhapvitrisach.Clear();
					}
					else
					{
						MessageBox.Show("Cập nhật vị trí sách không thành công!", "Thông báo",
							MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button_xoavitri_Click(object sender, EventArgs e)
		{
			try
			{
				if (dataGridView1.CurrentRow == null)
				{
					MessageBox.Show("Vui lòng chọn một vị trí sách để xóa!", "Thông báo",
						MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				string selectedBook = dataGridView1.CurrentRow.Cells["TenSach"].Value.ToString();

				DialogResult confirmResult = MessageBox.Show(
					$"Bạn có chắc chắn muốn xóa vị trí sách '{selectedBook}'?",
					"Xác nhận xóa", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

				if (confirmResult == DialogResult.Yes)
				{
					using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
					{
						sqlConnection.Open();
						string deleteQuery = "DELETE FROM vitrisach WHERE TenSach = @TenSach";
						SQLiteCommand deleteCommand = new SQLiteCommand(deleteQuery, sqlConnection);
						deleteCommand.Parameters.AddWithValue("@TenSach", selectedBook);

						int affectedRows = deleteCommand.ExecuteNonQuery();
						if (affectedRows > 0)
						{
							MessageBox.Show("Xóa vị trí sách thành công!", "Thông báo",
								MessageBoxButtons.OK, MessageBoxIcon.Information);

							LoadDataGridView();
							LoadBookLocations();
							textBox_nhapvitrisach.Clear();
						}
						else
						{
							MessageBox.Show("Xóa vị trí sách không thành công!", "Thông báo",
								MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi: " + ex.Message, "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void button_travitri_Click(object sender, EventArgs e)
		{
			try
			{
				string selectedBook = comboBox_travitrisach.Text.Trim();

				if (string.IsNullOrEmpty(selectedBook))
				{
					MessageBox.Show("Vui lòng chọn sách để tra cứu vị trí!", "Thông báo",
						MessageBoxButtons.OK, MessageBoxIcon.Warning);
					return;
				}

				using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
				{
					sqlConnection.Open();
					string query = "SELECT ViTri FROM vitrisach WHERE TenSach = @TenSach";
					SQLiteCommand command = new SQLiteCommand(query, sqlConnection);
					command.Parameters.AddWithValue("@TenSach", selectedBook);

					object result = command.ExecuteScalar();
					if (result != null)
					{
						string location = result.ToString();
						MessageBox.Show($"Vị trí sách là: {location}", "Thông tin vị trí",
							MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
					else
					{
						MessageBox.Show("Chưa cập nhập vị trí sách này, vui lòng cập nhập!",
							"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tra cứu vị trí sách: {ex.Message}", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void dataGridView1_SelectionChanged(object sender, EventArgs e)
		{
			try
			{
				if (dataGridView1.CurrentRow != null)
				{
					string selectedBook = dataGridView1.CurrentRow.Cells["TenSach"].Value.ToString();
					string location = dataGridView1.CurrentRow.Cells["ViTri"].Value.ToString();

					// Gán giá trị vào các điều khiển
					comboBox_chonsachthem.SelectedValue = selectedBook;
					textBox_nhapvitrisach.Text = location;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi chọn dữ liệu: {ex.Message}", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void comboBox_travitrisach_KeyDown(object sender, KeyEventArgs e)
		{
			try
			{
				if (e.KeyCode == Keys.Enter)
				{
					button_travitri_Click(sender, e);
					e.Handled = true;
					e.SuppressKeyPress = true;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi xử lý phím Enter: {ex.Message}", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void vitrisach_Load(object sender, EventArgs e)
		{
			try
			{
				LoadBookData();
				LoadBookLocations();
				LoadDataGridView();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tải form: {ex.Message}", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		// Tải dữ liệu sách từ bảng quanlysach
		private void LoadBookData()
		{
			try
			{
				DataTable dtBooks = modify.getAllbook();
				comboBox_chonsachthem.DisplayMember = "TenSach";
				comboBox_chonsachthem.ValueMember = "TenSach";
				comboBox_chonsachthem.DataSource = dtBooks;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tải dữ liệu sách: {ex.Message}", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		// Tải vị trí sách từ bảng vitrisach
		private void LoadBookLocations()
		{
			try
			{
				using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
				{
					sqlConnection.Open();
					string query = "SELECT TenSach FROM vitrisach";
					SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, sqlConnection);
					DataTable dt = new DataTable();
					adapter.Fill(dt);

					comboBox_travitrisach.DisplayMember = "TenSach";
					comboBox_travitrisach.ValueMember = "TenSach";
					comboBox_travitrisach.DataSource = dt;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tải danh sách vị trí sách: {ex.Message}", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		// Tải dữ liệu vào DataGridView
		private void LoadDataGridView()
		{
			try
			{
				using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
				{
					sqlConnection.Open();
					string query = "SELECT * FROM vitrisach";
					SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, sqlConnection);
					DataTable dt = new DataTable();
					adapter.Fill(dt);

					dataGridView1.DataSource = dt;
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tải dữ liệu vào bảng: {ex.Message}", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
