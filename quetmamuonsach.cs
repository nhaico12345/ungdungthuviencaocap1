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

namespace ungdungthuviencaocap
{
	public partial class quetmamuonsach : DevExpress.XtraEditors.XtraForm
	{
		FilterInfoCollection filterInfoCollection;
		VideoCaptureDevice videoCaptureDevice;

		public quetmamuonsach()
		{
			InitializeComponent();
			label_hienketqua.Text = "";
		}

		CancellationTokenSource cancellationToken;

		private void quetmamuonsach_Load(object sender, EventArgs e)
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
					button1.Text = "Dừng lại! ";
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

													// Truyền kết quả về form cha và đóng form
													if (Owner != null && Owner is muontrasach parentForm)
													{
														parentForm.SetBookCode(result.ToString());

														// Dừng camera
														if (videoCaptureDevice != null && videoCaptureDevice.IsRunning)
															videoCaptureDevice.Stop();

														// Hủy tác vụ quét
														cancellationToken.Cancel();

														// Đóng form
														this.BeginInvoke(new Action(() => {
															try
															{
																this.Close();
															}
															catch (Exception)
															{
																// Bỏ qua lỗi khi đóng form
															}
														}));
														return;
													}

													if (result.ResultPoints.Length > 0)
													{
														try
														{
															var offsetX = pictureBox1.SizeMode == PictureBoxSizeMode.Zoom
																? (pictureBox1.Width - pictureBox1.Image.Width) / 2
																: 0;
															var offsetY = pictureBox1.SizeMode == PictureBoxSizeMode.Zoom
																? (pictureBox1.Height - pictureBox1.Image.Height) / 2
																: 0;
															var rect = new Rectangle(
																(int)result.ResultPoints[0].X + offsetX,
																(int)result.ResultPoints[0].Y + offsetY,
																1, 1);

															foreach (var point in result.ResultPoints)
															{
																if (point.X + offsetX < rect.Left)
																	rect = new Rectangle((int)point.X + offsetX, rect.Y, rect.Width + rect.X - (int)point.X - offsetX, rect.Height);
																if (point.X + offsetX > rect.Right)
																	rect = new Rectangle(rect.X, rect.Y, rect.Width + (int)point.X - (rect.X - offsetX), rect.Height);
																if (point.Y + offsetY < rect.Top)
																	rect = new Rectangle(rect.X, (int)point.Y + offsetY, rect.Width, rect.Height + rect.Y - (int)point.Y - offsetY);
																if (point.Y + offsetY > rect.Bottom)
																	rect = new Rectangle(rect.X, rect.Y, rect.Width, rect.Height + (int)point.Y - (rect.Y - offsetY));
															}

															using (var g = pictureBox1.CreateGraphics())
															{
																using (Pen pen = new Pen(Color.Green, 5))
																{
																	g.DrawRectangle(pen, rect);
																	pen.Color = Color.Yellow;
																	pen.DashPattern = new float[] { 5, 5 };
																	g.DrawRectangle(pen, rect);
																}
																g.DrawString(result.ToString(), new Font("Tahoma", 16f), Brushes.Blue, new Point(rect.X - 60, rect.Y - 50));
															}
														}
														catch (Exception)
														{
															// Bỏ qua lỗi vẽ hình chữ nhật
														}
													}
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

		private new void FormClosing(object sender, FormClosingEventArgs e)
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
	}
}
