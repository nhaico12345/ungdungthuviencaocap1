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
using System.Reflection;
using System.IO;

namespace ungdungthuviencaocap
{
	public partial class thongtintaikhoan : DevExpress.XtraEditors.XtraForm
	{
		public thongtintaikhoan()
		{
			InitializeComponent();
			LoadUserInfo();
		}
		private void LoadUserInfo()
		{
			try
			{
				// Lấy thông tin người dùng đã đăng nhập
				var currentUser = UserSession.CurrentUser;
				if (currentUser != null)
				{
					// Điền các trường văn bản với thông tin người dùng
					textBox_hovaten.Text = currentUser.HoVaTen;
					textBox_ngaysinh.Text = currentUser.NgaySinh;
					textBox_masv.Text = currentUser.MaSinhVien;
					textBox_gioitinh.Text = currentUser.GioiTinh;
					textbox_diachi.Text = currentUser.DiaChi;
					textBox_tendangnhap.Text = currentUser.TenTaiKhoan;
					textBox_email.Text = currentUser.Email;
					textBox_quyen.Text = currentUser.Quyen;
					textBox_sodienthoai.Text = currentUser.SoDienThoai;

					// Tải avatar của người dùng sử dụng đường dẫn tương đối
					if (!string.IsNullOrEmpty(currentUser.Avatar))
					{
						try
						{
							// Sử dụng đường dẫn tương đối đến thư mục Avatar
							string avatarFolder = Path.Combine(Application.StartupPath, "Avatar");
							string avatarPath = Path.Combine(avatarFolder, currentUser.Avatar);

							if (File.Exists(avatarPath))
							{
								// Tải và hiển thị hình ảnh
								using (FileStream fs = new FileStream(avatarPath, FileMode.Open, FileAccess.Read))
								{
									pictureBox_hinhanh.Image = Image.FromStream(fs);
									pictureBox_hinhanh.SizeMode = PictureBoxSizeMode.StretchImage;
								}
							}
							else
							{
								// Nếu không tìm thấy file ảnh cụ thể, tải ảnh mặc định
								string defaultImagePath = Path.Combine(avatarFolder, "default.png");
								if (File.Exists(defaultImagePath))
								{
									using (FileStream fs = new FileStream(defaultImagePath, FileMode.Open, FileAccess.Read))
									{
										pictureBox_hinhanh.Image = Image.FromStream(fs);
										pictureBox_hinhanh.SizeMode = PictureBoxSizeMode.StretchImage;
									}
								}
								else
								{
									MessageBox.Show("Không tìm thấy ảnh của người dùng và ảnh mặc định. ",
										"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
								}
							}
						}
						catch (Exception ex)
						{
							MessageBox.Show("Không thể tải ảnh: " + ex.Message, "Lỗi",
								MessageBoxButtons.OK, MessageBoxIcon.Error);
						}
					}
				}
				else
				{
					MessageBox.Show("Không có dữ liệu người dùng. ", "Thông báo",
						MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Có lỗi khi tải thông tin người dùng: " + ex.Message,
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void thongtintaikhoan_Load(object sender, EventArgs e)
		{

		}

		private void button_thayanh_Click(object sender, EventArgs e)
		{
			try
			{
				using (OpenFileDialog openFileDialog = new OpenFileDialog())
				{
					openFileDialog.Filter = "Tệp hình ảnh|*.jpg;*.jpeg;*.png;*.gif;*.bmp";
					openFileDialog.Title = "Chọn ảnh đại diện";
					if (openFileDialog.ShowDialog() == DialogResult.OK)
					{
						// Đường dẫn tương đối đến thư mục avatar
						string avatarFolder = Path.Combine(Application.StartupPath, "Avatar");

						// Đảm bảo thư mục tồn tại
						if (!Directory.Exists(avatarFolder))
						{
							Directory.CreateDirectory(avatarFolder);
						}

						// Tạo tên file mới dựa trên tên đăng nhập và thời gian
						string newFileName = UserSession.CurrentUser.TenTaiKhoan + "_" +
							DateTime.Now.ToString("yyyyMMddHHmmss") +
							Path.GetExtension(openFileDialog.FileName);

						// Đường dẫn đầy đủ của file mới
						string newFilePath = Path.Combine(avatarFolder, newFileName);

						// Sao chép file ảnh đã chọn vào thư mục Avatar
						File.Copy(openFileDialog.FileName, newFilePath);

						// Hiển thị ảnh mới
						pictureBox_hinhanh.Image = Image.FromFile(newFilePath);
						pictureBox_hinhanh.SizeMode = PictureBoxSizeMode.StretchImage;

						// Cập nhật tên ảnh trong CSDL
						UpdateAvatarInDatabase(UserSession.CurrentUser.TenTaiKhoan, newFileName);

						// Cập nhật thông tin người dùng hiện tại
						UserSession.CurrentUser.Avatar = newFileName;

						MessageBox.Show("Đã thay đổi ảnh đại diện thành công! ", "Thông báo",
							MessageBoxButtons.OK, MessageBoxIcon.Information);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Có lỗi khi thay đổi ảnh: " + ex.Message, "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}


		private void UpdateAvatarInDatabase(string tenTaiKhoan, string avatarFileName)
		{
			try
			{
				string query = "UPDATE TaiKhoan SET Avatar = @AvatarFileName WHERE TenTaiKhoan = @TenTaiKhoan";

				Dictionary<string, object> parameters = new Dictionary<string, object>
		{
			{ "@AvatarFileName", avatarFileName },
			{ "@TenTaiKhoan", tenTaiKhoan }
		};

				int result = Connection.ExecuteNonQuery(query, parameters);

				if (result <= 0)
				{
					MessageBox.Show("Không thể cập nhật ảnh đại diện trong cơ sở dữ liệu.",
						"Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi cập nhật cơ sở dữ liệu: " + ex.Message,
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
		// Biến để theo dõi trạng thái chỉnh sửa
		private bool isEditing = false;
		private void button_capnhatthongtin_Click(object sender, EventArgs e)
		{
			if (!isEditing) // Chế độ xem - chuyển sang chế độ chỉnh sửa
			{
				// Hiển thị thông báo xác nhận
				DialogResult result = MessageBox.Show("Bạn muốn cập nhật thông tin?",
					"Xác nhận", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

				if (result == DialogResult.Yes)
				{
					// Chuyển các textbox sang chế độ chỉnh sửa (bỏ readonly)
					textBox_hovaten.ReadOnly = false;
					textbox_diachi.ReadOnly = false;
					textBox_email.ReadOnly = false;
					textBox_ngaysinh.ReadOnly = false;
					textBox_gioitinh.ReadOnly = false;
					textBox_sodienthoai.ReadOnly = false;

					// Đổi text của button
					button_capnhatthongtin.Text = "     Lưu thông tin";

					// Đánh dấu đang ở chế độ chỉnh sửa
					isEditing = true;
				}
			}
			else // Chế độ chỉnh sửa - lưu thông tin và chuyển về chế độ xem
			{
				try
				{
					// Kiểm tra dữ liệu nhập vào
					if (string.IsNullOrWhiteSpace(textBox_hovaten.Text))
					{
						MessageBox.Show("Họ và tên không được để trống!", "Lỗi",
							MessageBoxButtons.OK, MessageBoxIcon.Error);
						return;
					}

					// Cập nhật thông tin vào CSDL
					string query = @"UPDATE TaiKhoan 
                             SET HoVaTen = @HoVaTen, 
                                 DiaChi = @DiaChi, 
                                 Email = @Email,
                                 NgaySinh = @NgaySinh,
                                 GioiTinh = @GioiTinh,
								 SoDienThoai = @SoDienThoai
                             WHERE TenTaiKhoan = @TenTaiKhoan";

					Dictionary<string, object> parameters = new Dictionary<string, object>
			{
				{ "@HoVaTen", textBox_hovaten.Text },
				{ "@DiaChi", textbox_diachi.Text },
				{ "@Email", textBox_email.Text },
				{ "@NgaySinh", textBox_ngaysinh.Text },
				{ "@GioiTinh", textBox_gioitinh.Text },
				{ "@SoDienThoai", textBox_sodienthoai.Text },
				{ "@TenTaiKhoan", UserSession.CurrentUser.TenTaiKhoan }
			};

					int result = Connection.ExecuteNonQuery(query, parameters);

					if (result > 0)
					{
						// Cập nhật thông tin người dùng hiện tại
						UserSession.CurrentUser.HoVaTen = textBox_hovaten.Text;
						UserSession.CurrentUser.DiaChi = textbox_diachi.Text;
						UserSession.CurrentUser.Email = textBox_email.Text;
						UserSession.CurrentUser.NgaySinh = textBox_ngaysinh.Text;
						UserSession.CurrentUser.GioiTinh = textBox_gioitinh.Text;
						UserSession.CurrentUser.SoDienThoai = textBox_sodienthoai.Text;
						MessageBox.Show("Cập nhật thông tin thành công!", "Thông báo",
							MessageBoxButtons.OK, MessageBoxIcon.Information);

						// Đặt các textbox về chế độ chỉ đọc
						textBox_hovaten.ReadOnly = true;
						textbox_diachi.ReadOnly = true;
						textBox_email.ReadOnly = true;
						textBox_ngaysinh.ReadOnly = true;
						textBox_gioitinh.ReadOnly = true;
						textBox_sodienthoai.ReadOnly = true;
						// Đổi text của button về trạng thái ban đầu
						button_capnhatthongtin.Text = "     Cập nhật thông tin";

						// Đánh dấu đã thoát chế độ chỉnh sửa
						isEditing = false;
					}
					else
					{
						MessageBox.Show("Không thể cập nhật thông tin!", "Thông báo",
							MessageBoxButtons.OK, MessageBoxIcon.Warning);
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show("Lỗi khi cập nhật thông tin: " + ex.Message, "Lỗi",
						MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
		}
	}
}