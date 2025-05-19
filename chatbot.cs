using System;
using System.Windows.Forms;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using DevExpress.XtraEditors;

namespace ungdungthuviencaocap
{
	public partial class chatbot : XtraForm
	{
		// Tạo HttpClient tĩnh được sử dụng lại cho tất cả các yêu cầu
		private static readonly HttpClient client = new HttpClient();

		// Thay thế bằng khóa API Google AI Studio của bạn
		private const string API_KEY = "AIzaSyA5iWWIcLvvkB850REGCros3Jq58coBMlw";

		// Đường dẫn API Gemini 2.0 Flash
		private const string API_URL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent";

		public chatbot()
		{
			InitializeComponent();

			// Cấu hình HTTP client với các header cần thiết
			ConfigureHttpClient();

			// Hiển thị tin nhắn chào mừng
			richTextBox_hienthitinnhan.AppendText("Trợ lý thư viện: Chào bạn tôi là trợ lý thư viện, bạn có thể hỏi tôi bất cứ câu hỏi nào.\n\n");
		}

		private void ConfigureHttpClient()
		{
			// Thiết lập header mặc định cho HTTP client
			client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
		}

		private async void button_gui_Click(object sender, EventArgs e)
		{
			string userMessage = textBox_nhapcauhoi.Text.Trim();

			if (!string.IsNullOrWhiteSpace(userMessage))
			{
				// Hiển thị tin nhắn người dùng
				richTextBox_hienthitinnhan.AppendText($"Bạn: {userMessage}\n\n");

				// Xóa nội dung trong ô nhập liệu
				textBox_nhapcauhoi.Clear();

				// Đặt con trỏ chuột ở trạng thái đang chờ
				this.Cursor = Cursors.WaitCursor;

				try
				{
					// Lấy phản hồi từ AI
					string botResponse = await GetAIResponse(userMessage);

					// Hiển thị phản hồi của bot
					richTextBox_hienthitinnhan.AppendText($"Trợ lý thư viện: {botResponse}\n\n");
				}
				catch (Exception ex)
				{
					richTextBox_hienthitinnhan.AppendText($"Trợ lý thư viện: Xin lỗi, có lỗi xảy ra: {ex.Message}\n\n");
				}
				finally
				{
					// Đặt lại con trỏ chuột
					this.Cursor = Cursors.Default;

					// Cuộn xuống cuối của rich text box
					richTextBox_hienthitinnhan.ScrollToCaret();
				}
			}
		}

		private async Task<string> GetAIResponse(string userMessage)
		{
			try
			{
				// Tạo URL yêu cầu với khóa API
				string requestUrl = $"{API_URL}?key={API_KEY}";

				// Tạo nội dung yêu cầu theo đúng định dạng API Gemini
				var requestBody = new
				{
					contents = new[]
					{
						new
						{
							parts = new[]
							{
								new { text = userMessage }
							}
						}
					}
				};

				// Chuyển đổi nội dung yêu cầu sang JSON
				string jsonRequest = JsonConvert.SerializeObject(requestBody);

				// Tạo nội dung HTTP
				var content = new StringContent(jsonRequest, Encoding.UTF8, "application/json");

				// Gửi yêu cầu
				var response = await client.PostAsync(requestUrl, content);

				// Đảm bảo yêu cầu thành công
				response.EnsureSuccessStatusCode();

				// Đọc nội dung phản hồi
				string jsonResponse = await response.Content.ReadAsStringAsync();

				// Giải mã phản hồi
				dynamic responseObject = JsonConvert.DeserializeObject(jsonResponse);

				// Trích xuất nội dung văn bản từ phản hồi
				string aiResponse = responseObject.candidates[0].content.parts[0].text;

				return aiResponse;
			}
			catch (HttpRequestException ex)
			{
				Console.WriteLine($"Lỗi HTTP khi gọi API Gemini: {ex.Message}");
				return "Xin lỗi, không thể kết nối đến dịch vụ AI. Vui lòng kiểm tra kết nối mạng và API key.";
			}
			catch (JsonException ex)
			{
				Console.WriteLine($"Lỗi phân tích JSON: {ex.Message}");
				return "Xin lỗi, có lỗi xảy ra khi xử lý phản hồi từ AI.";
			}
			catch (Exception ex)
			{
				Console.WriteLine($"Lỗi không xác định: {ex.Message}");
				return "Xin lỗi, tôi không thể xử lý yêu cầu của bạn lúc này. Vui lòng thử lại sau.";
			}
		}

		private void textBox_nhapcauhoi_KeyDown(object sender, KeyEventArgs e)
		{
			// Kích hoạt nút gửi khi phím Enter được nhấn
			if (e.KeyCode == Keys.Enter)
			{
				button_gui_Click(sender, e);
				e.Handled = true;
				e.SuppressKeyPress = true;
			}
		}
	}
}
