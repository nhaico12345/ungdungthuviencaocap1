using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Data.SQLite;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Newtonsoft.Json;
using System.Net;
using System.Diagnostics;

namespace ungdungthuviencaocap
{
	public partial class caidat : DevExpress.XtraEditors.XtraForm
	{
		private trangchu _mainForm;
		private const string DatabaseFileName = "ungdungthuvien.db";

		// Thông tin đăng nhập Yandex Disk
		private const string YandexUsername = "oosp0305@gmail.com";
		private const string YandexPassword = "inxkjzdopljnngzz";
		private const string YandexWebDavUrl = "https://webdav.yandex.ru";

		private static SQLiteConnection DbConnection;
		public caidat()
		{
			InitializeComponent();

			// Kết nối sự kiện nhấn nút
			this.button_thaymau.Click += new EventHandler(button_thaymau_Click);

			// Đính kèm một hàm xử lý sự kiện cho sự kiện Load để tìm form chính
			this.Load += (sender, e) => FindMainForm();
		}

		private void FindMainForm()
		{
			// Cố gắng tìm form chính thông qua cây điều khiển
			Control parent = this.Parent;
			while (parent != null && !(parent is trangchu))
			{
				parent = parent.Parent;
			}

			if (parent != null && parent is trangchu)
			{
				_mainForm = parent as trangchu;
			}
			else
			{
				// Phương án thay thế: tìm trong danh sách các form đang mở
				foreach (Form form in Application.OpenForms)
				{
					if (form is trangchu)
					{
						_mainForm = form as trangchu;
						break;
					}
				}
			}
		}

		private void button_thaymau_Click(object sender, EventArgs e)
		{
			try
			{
				// Đảm bảo chúng ta có tham chiếu đến form chính
				if (_mainForm == null)
				{
					FindMainForm();
				}

				if (_mainForm != null)
				{
					// Lấy màu đã chọn từ color picker
					System.Drawing.Color selectedColor = (System.Drawing.Color)colorPickEdit1.EditValue;

					// Thay đổi màu cho control accordion thông qua phương thức công khai
					_mainForm.ChangeAccordionControlColor(selectedColor);

					MessageBox.Show("Màu giao diện đã được thay đổi thành công!", "Thông báo",
						MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				else
				{
					MessageBox.Show("Không thể tìm thấy form chính để thay đổi màu giao diện.", "Lỗi",
						MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi thay đổi màu giao diện: {ex.Message}", "Lỗi",
					MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		// Phương thức tải lên Yandex Disk qua WebDAV
		private void button_dongbolen_Click(object sender, EventArgs e)
		{
			DialogResult result = MessageBox.Show(
				"Bạn muốn đồng bộ dữ liệu lên server?",
				"Xác nhận",
				MessageBoxButtons.YesNo,
				MessageBoxIcon.Question);

			if (result == DialogResult.Yes)
			{
				Cursor = Cursors.WaitCursor;

				try
				{
					bool success = UploadToYandexDiskWebDav();

					if (success)
					{
						MessageBox.Show(
							"Đồng bộ dữ liệu lên server thành công!",
							"Thông báo",
							MessageBoxButtons.OK,
							MessageBoxIcon.Information);
					}
					else
					{
						MessageBox.Show(
							"Đồng bộ dữ liệu lên server thất bại!",
							"Lỗi",
							MessageBoxButtons.OK,
							MessageBoxIcon.Error);
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show(
						$"Đồng bộ dữ liệu lên server thất bại! Lỗi: {ex.Message}",
						"Lỗi",
						MessageBoxButtons.OK,
						MessageBoxIcon.Error);
				}
				finally
				{
					Cursor = Cursors.Default;
				}
			}
		}

		private void button_dongboxuong_Click(object sender, EventArgs e)
		{
			DialogResult result = MessageBox.Show(
				"Bạn muốn đồng bộ dữ liệu từ server xuống, dữ liệu từ server sẽ ghi đè lên dữ liệu cũ bạn vẫn muốn tiếp tục chứ?",
				"Cảnh báo",
				MessageBoxButtons.YesNo,
				MessageBoxIcon.Warning);

			if (result == DialogResult.Yes)
			{
				Cursor = Cursors.WaitCursor;

				try
				{
					// Đóng tất cả kết nối đến cơ sở dữ liệu
					CloseAllDatabaseConnections();

					// Chờ một chút để đảm bảo kết nối được đóng hoàn toàn
					System.Threading.Thread.Sleep(1000);

					bool success = DownloadFromYandexDiskWebDav();

					if (success)
					{
						MessageBox.Show(
							"Đồng bộ dữ liệu từ server xuống thành công!",
							"Thông báo",
							MessageBoxButtons.OK,
							MessageBoxIcon.Information);

						// Khởi động lại form trang chủ để tải lại dữ liệu mới (tùy chọn)
						if (_mainForm != null && MessageBox.Show(
							"Dữ liệu đã được cập nhật. Bạn có muốn khởi động lại ứng dụng để áp dụng thay đổi không?",
							"Khởi động lại",
							MessageBoxButtons.YesNo,
							MessageBoxIcon.Question) == DialogResult.Yes)
						{
							Application.Restart();
						}
					}
					else
					{
						MessageBox.Show(
							"Đồng bộ dữ liệu từ server xuống thất bại!",
							"Lỗi",
							MessageBoxButtons.OK,
							MessageBoxIcon.Error);
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show(
						$"Đồng bộ dữ liệu từ server xuống thất bại! Lỗi: {ex.Message}",
						"Lỗi",
						MessageBoxButtons.OK,
						MessageBoxIcon.Error);
				}
				finally
				{
					Cursor = Cursors.Default;
				}
			}
		}

		// Đóng tất cả kết nối đến cơ sở dữ liệu
		private void CloseAllDatabaseConnections()
		{
			try
			{
				// Nếu đang sử dụng SQLite và có biến kết nối tĩnh
				if (DbConnection != null && DbConnection.State != System.Data.ConnectionState.Closed)
				{
					DbConnection.Close();
					DbConnection.Dispose();
					DbConnection = null;
				}

				// Thông báo GC thu gom rác
				GC.Collect();
				GC.WaitForPendingFinalizers();

				// Nếu cần, có thể thực hiện các bước để đóng kết nối cụ thể cho ứng dụng ở đây

				// Nếu cần, thử giải phóng file bằng cách gọi đến tất cả form
				if (_mainForm != null)
				{
					// Gọi phương thức để đóng kết nối cơ sở dữ liệu nếu bạn có một trong trang chủ
					// _mainForm.CloseDatabase();
				}
			}
			catch (Exception ex)
			{
				Debug.WriteLine($"Lỗi khi đóng kết nối cơ sở dữ liệu: {ex.Message}");
			}
		}

		// Phương thức tải xuống sử dụng file tạm
		private bool DownloadFromYandexDiskWebDav()
		{
			string localFilePath = Path.Combine(Application.StartupPath, DatabaseFileName);
			string tempFilePath = Path.Combine(Application.StartupPath, $"temp_{DatabaseFileName}");

			try
			{
				// Tạo request đến WebDAV
				string remoteUrl = $"{YandexWebDavUrl}/{DatabaseFileName}";
				HttpWebRequest request = (HttpWebRequest)WebRequest.Create(remoteUrl);
				request.Method = "GET";
				request.Credentials = new NetworkCredential(YandexUsername, YandexPassword);

				// Đầu tiên tải về file tạm
				using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
				{
					using (Stream responseStream = response.GetResponseStream())
					{
						using (FileStream fileStream = File.Create(tempFilePath))
						{
							byte[] buffer = new byte[8192];
							int bytesRead;
							while ((bytesRead = responseStream.Read(buffer, 0, buffer.Length)) > 0)
							{
								fileStream.Write(buffer, 0, bytesRead);
							}
						}
					}

					if (response.StatusCode != HttpStatusCode.OK)
					{
						return false;
					}
				}

				// Kiểm tra xem file tạm có tồn tại và hợp lệ không
				if (!File.Exists(tempFilePath) || new FileInfo(tempFilePath).Length == 0)
				{
					throw new Exception("File tải xuống không hợp lệ hoặc rỗng");
				}

				// Thử xóa file cũ nếu tồn tại
				if (File.Exists(localFilePath))
				{
					try
					{
						File.Delete(localFilePath);
					}
					catch (IOException)
					{
						// Nếu không thể xóa file, thử giải phóng nó trước
						CloseAllDatabaseConnections();
						System.Threading.Thread.Sleep(1000); // Chờ một chút

						// Thử lại việc xóa
						File.Delete(localFilePath);
					}
				}

				// Di chuyển file tạm để thay thế file gốc
				File.Move(tempFilePath, localFilePath);

				return true;
			}
			catch (WebException ex)
			{
				if (ex.Response != null)
				{
					using (StreamReader reader = new StreamReader(ex.Response.GetResponseStream()))
					{
						string errorResponse = reader.ReadToEnd();
						throw new Exception($"Lỗi WebDAV: {errorResponse}");
					}
				}
				throw;
			}
			finally
			{
				// Dọn dẹp file tạm nếu còn tồn tại
				if (File.Exists(tempFilePath))
				{
					try
					{
						File.Delete(tempFilePath);
					}
					catch
					{
						// Bỏ qua lỗi khi xóa file tạm
					}
				}
			}
		}

		// Phương thức tải lên sử dụng WebDAV
		private bool UploadToYandexDiskWebDav()
		{
			string localFilePath = Path.Combine(Application.StartupPath, DatabaseFileName);

			if (!File.Exists(localFilePath))
			{
				throw new FileNotFoundException($"File {DatabaseFileName} không tồn tại.");
			}

			try
			{
				// Tạo request đến WebDAV
				string remoteUrl = $"{YandexWebDavUrl}/{DatabaseFileName}";
				HttpWebRequest request = (HttpWebRequest)WebRequest.Create(remoteUrl);
				request.Method = "PUT";
				request.Credentials = new NetworkCredential(YandexUsername, YandexPassword);

				// Thiết lập nội dung của request
				byte[] fileContents = File.ReadAllBytes(localFilePath);
				request.ContentLength = fileContents.Length;

				// Gửi dữ liệu
				using (Stream requestStream = request.GetRequestStream())
				{
					requestStream.Write(fileContents, 0, fileContents.Length);
				}

				// Lấy phản hồi
				using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
				{
					return response.StatusCode == HttpStatusCode.Created ||
						   response.StatusCode == HttpStatusCode.OK ||
						   response.StatusCode == HttpStatusCode.NoContent;
				}
			}
			catch (WebException ex)
			{
				if (ex.Response != null)
				{
					using (StreamReader reader = new StreamReader(ex.Response.GetResponseStream()))
					{
						string errorResponse = reader.ReadToEnd();
						throw new Exception($"Lỗi WebDAV: {errorResponse}");
					}
				}
				throw;
			}
		}

		private void button_kiemtraketnoi_Click(object sender, EventArgs e)
		{
			try
			{
				using (SQLiteConnection conn = Connection.GetSQLiteConnection())
				{
					conn.Open();
					MessageBox.Show("Kết nối thành công", "Thông báo");
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Kết nối thất bại: ", "Thông báo" + ex.Message);
			}
		}
	}
}

