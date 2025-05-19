using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.IO;
using System.Data;
using System.Reflection;

namespace ungdungthuviencaocap
{
	class Connection
	{
		private static string GetConnectionString()
		{
			try
			{
				// Lấy đường dẫn của thư mục chứa ứng dụng đang chạy
				string appPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

				// Định nghĩa đường dẫn tương đối đến file database
				string dbPath = Path.Combine(appPath, "ungdungthuvien.db");

				// Kiểm tra xem file có tồn tại không
				if (!File.Exists(dbPath))
				{
					throw new FileNotFoundException($"Không tìm thấy file database tại: {dbPath}");
				}

				return $"Data Source={dbPath};Version=3;";
			}
			catch (Exception ex)
			{
				// Xử lý ngoại lệ - hiển thị thông báo hoặc ghi log
				Console.WriteLine($"Lỗi khi tạo kết nối: {ex.Message}");
				throw; // Ném lại ngoại lệ để xử lý ở lớp gọi
			}
		}

		public static SQLiteConnection GetSQLiteConnection()
		{
			return new SQLiteConnection(GetConnectionString());
		}

		public static DataTable GetData(string query)
		{
			DataTable dt = new DataTable();
			using (SQLiteConnection conn = GetSQLiteConnection())
			{
				conn.Open();
				using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
				{
					using (SQLiteDataAdapter adapter = new SQLiteDataAdapter(cmd))
					{
						adapter.Fill(dt);
					}
				}
			}
			return dt;
		}

		// Phương thức thực thi câu lệnh không truy vấn (INSERT, UPDATE, DELETE)
		public static int ExecuteNonQuery(string query, Dictionary<string, object> parameters = null)
		{
			using (SQLiteConnection conn = GetSQLiteConnection())
			{
				conn.Open();
				using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
				{
					if (parameters != null)
					{
						foreach (var param in parameters)
						{
							cmd.Parameters.AddWithValue(param.Key, param.Value);
						}
					}
					return cmd.ExecuteNonQuery();
				}
			}
		}
	}
}
