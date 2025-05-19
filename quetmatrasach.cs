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
using AForge.Video.DirectShow;
using AForge.Video;
using System.Media;
using System.Reflection;
using System.Threading;
using ZXing;
using System.Data.SQLite;

namespace ungdungthuviencaocap
{
	public partial class quetmatrasach : DevExpress.XtraEditors.XtraForm
	{
		FilterInfoCollection filterInfoCollection;
		VideoCaptureDevice videoCaptureDevice;
		public quetmatrasach()
		{
			InitializeComponent();
			this.FormClosing += Form_Closing;
			label_hienketqua.Text = "";
		}

		CancellationTokenSource cancellationToken;

		private void button1_Click(object sender, EventArgs e)
		{
			try
			{
				if (button1.Text == "Bắt đầu")
				{
					// Kiểm tra xem có camera được chọn không
					if (comboBox_camera.SelectedIndex < 0)
					{
						MessageBox.Show("Vui lòng chọn camera trước khi bắt đầu!", "Thông báo",
							MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return;
					}

					// Sử dụng camera được chọn từ comboBox_camera
					videoCaptureDevice = new VideoCaptureDevice(
						filterInfoCollection[comboBox_camera.SelectedIndex].MonikerString);
					videoCaptureDevice.NewFrame += FinalFrame_NewFrame;
					videoCaptureDevice.Start();

					cancellationToken = new CancellationTokenSource();
					var sourcetoken = cancellationToken.Token;
					onStartScan(sourcetoken);

					button1.Text = "Dừng lại!";
				}
				else
				{
					button1.Text = "Bắt đầu";
					// Dừng quét và giải phóng tài nguyên
					try
					{
						cancellationToken.Cancel();
						if (videoCaptureDevice != null && videoCaptureDevice.IsRunning)
							videoCaptureDevice.Stop();
					}
					catch (Exception ex)
					{
						MessageBox.Show($"Lỗi khi dừng camera: {ex.Message}",
							"Cảnh báo", MessageBoxButtons.OK, MessageBoxIcon.Warning);
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Đã xảy ra lỗi: {ex.Message}",
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				// Đảm bảo nút Bắt đầu được đặt lại nếu có lỗi
				button1.Text = "Bắt đầu";
			}
		}

		private void FinalFrame_NewFrame(object sender, NewFrameEventArgs eventArgs)
		{
			try
			{
				pictureBox1.Image = (Bitmap)eventArgs.Frame.Clone();
			}
			catch (Exception)
			{
				// Bỏ qua lỗi khi xử lý frame (có thể xảy ra khi đóng form)
			}
		}


		public void onStartScan(CancellationToken sourcetoken)
		{
			Task.Factory.StartNew(new Action(() =>
			{
				while (true)
				{
					if (sourcetoken.IsCancellationRequested)
					{
						return;
					}

					Thread.Sleep(50);
					BarcodeReader Reader = new BarcodeReader();
					try
					{
						pictureBox1.BeginInvoke(new Action(() =>
						{
							try
							{
								if (pictureBox1.Image != null)
								{
									try
									{
										var results = Reader.DecodeMultiple((Bitmap)pictureBox1.Image);
										if (results != null)
										{
											foreach (Result result in results)
											{
												try
												{
													label_hienketqua.Text = result.ToString() + $"- Type: {result.BarcodeFormat.ToString()}";
													SystemSounds.Beep.Play();

													// Lấy mã đã quét được
													string maQuet = result.ToString();

													// Xử lý mã quét được
													ProcessScannedCode(maQuet);
													return; // Kết thúc sau khi xử lý mã đầu tiên
												}
												catch (Exception)
												{
													// Bỏ qua lỗi khi xử lý kết quả riêng lẻ
												}
											}
										}
									}
									catch (Exception)
									{
										// Bỏ qua lỗi khi giải mã
									}
								}
							}
							catch (ObjectDisposedException)
							{
								// Bỏ qua khi đối tượng đã bị hủy (có thể xảy ra khi đóng form)
								sourcetoken.ThrowIfCancellationRequested();
							}
							catch (Exception)
							{
								// Bỏ qua các ngoại lệ khác trong quá trình xử lý hình ảnh
							}
						}));
					}
					catch (ObjectDisposedException)
					{
						// Form đã bị đóng hoặc đối tượng đã bị hủy
						return;
					}
					catch (Exception)
					{
						// Bỏ qua các lỗi khác trong quá trình quét
					}
				}
			}), sourcetoken);
		}

		// Thêm phương thức xử lý mã đã quét
		private void ProcessScannedCode(string maQuet)
		{
			try
			{
				// Dừng camera
				if (videoCaptureDevice != null && videoCaptureDevice.IsRunning)
					videoCaptureDevice.Stop();

				// Hủy tác vụ quét
				if (cancellationToken != null && !cancellationToken.IsCancellationRequested)
					cancellationToken.Cancel();

				// Tìm kiếm thông tin sinh viên và sách từ cơ sở dữ liệu
				string maSinhVien = "";
				string hoVaTen = "";
				bool timThay = false;

				using (SQLiteConnection con = Connection.GetSQLiteConnection())
				{
					con.Open();
					string query = "SELECT MaSinhVien, HoVaTen FROM muontrasach WHERE MaSach = @MaSach AND IsActive = 1";
					SQLiteCommand command = new SQLiteCommand(query, con);
					command.Parameters.AddWithValue("@MaSach", maQuet);

					using (SQLiteDataReader reader = command.ExecuteReader())
					{
						if (reader.Read())
						{
							maSinhVien = reader["MaSinhVien"].ToString();
							hoVaTen = reader["HoVaTen"].ToString();
							timThay = true;
						}
					}
				}

				// Đóng form và truyền dữ liệu quét về form cha (muontrasach)
				this.BeginInvoke(new Action(() => {
					try
					{
						if (Owner != null && Owner is muontrasach parentForm)
						{
							if (timThay)
							{
								// Đóng form và gọi phương thức xử lý từ form cha
								this.Close();
								parentForm.XuLyMaSachQuet(maQuet, maSinhVien, hoVaTen);
							}
							else
							{
								// Không tìm thấy sách
								this.Close();
								MessageBox.Show($"Không tìm thấy sách có mã {maQuet} trong danh sách mượn.",
									"Không tìm thấy", MessageBoxButtons.OK, MessageBoxIcon.Information);
							}
						}
					}
					catch (Exception ex)
					{
						MessageBox.Show($"Lỗi khi xử lý mã sách: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
						this.Close();
					}
				}));
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi xử lý mã quét: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				this.Close();
			}
		}

		private void Form_Closing(object sender, FormClosingEventArgs e)
		{
			try
			{
				if (cancellationToken != null && !cancellationToken.IsCancellationRequested)
				{
					cancellationToken.Cancel();
				}

				if (videoCaptureDevice != null && videoCaptureDevice.IsRunning)
				{
					videoCaptureDevice.NewFrame -= FinalFrame_NewFrame;
					videoCaptureDevice.SignalToStop();
					videoCaptureDevice.WaitForStop();
					videoCaptureDevice = null;
				}
			}
			catch (Exception ex)
			{
				// Ghi log lỗi nếu cần thiết, nhưng không hiển thị MessageBox vì đang đóng form
				Console.WriteLine($"Lỗi khi đóng form: {ex.Message}");
			}
		}

		private void quetmatrasach_Load(object sender, EventArgs e)
		{
			try
			{
				// Sử dụng comboBox_camera có sẵn từ designer
				comboBox_camera.Items.Clear();
				filterInfoCollection = new FilterInfoCollection(FilterCategory.VideoInputDevice);

				// Thêm tất cả thiết bị camera được phát hiện vào comboBox_camera
				foreach (FilterInfo Device in filterInfoCollection)
					comboBox_camera.Items.Add(Device.Name);

				// Nếu có ít nhất một camera, chọn camera đầu tiên làm mặc định
				if (comboBox_camera.Items.Count > 0)
				{
					comboBox_camera.SelectedIndex = 0;
					videoCaptureDevice = new VideoCaptureDevice();
				}
				else
				{
					MessageBox.Show("Không tìm thấy thiết bị camera nào! Vui lòng kết nối camera và khởi động lại ứng dụng.",
						"Không tìm thấy camera", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Đã xảy ra lỗi khi khởi tạo camera: {ex.Message}",
					"Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
