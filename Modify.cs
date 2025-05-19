using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using System.Windows.Forms;
using QRCoder;

namespace ungdungthuviencaocap
{
	class Modify
	{
		SQLiteDataAdapter dataAdapter;
		SQLiteCommand sqlcommand;
		private static Random randomGenerator = new Random();

		public Modify()
		{
		}

		public string GenerateRandomAlphaNumericCode(int length = 8, string prefix = "")
		{
			const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
			string randomPart = new string(Enumerable.Repeat(chars, length)
			  .Select(s => s[randomGenerator.Next(s.Length)]).ToArray());
			return prefix + randomPart;
		}

		private bool IsMaTacGiaUnique(string maTacGia, SQLiteConnection connection)
		{
			string checkQuery = "SELECT COUNT(*) FROM tacgia WHERE MaTacGia = @MaTacGia COLLATE NOCASE";
			using (SQLiteCommand checkCmd = new SQLiteCommand(checkQuery, connection))
			{
				checkCmd.Parameters.AddWithValue("@MaTacGia", maTacGia);
				long count = (long)checkCmd.ExecuteScalar();
				return count == 0;
			}
		}
		private bool IsTenTacGiaUnique(string hoTen, SQLiteConnection connection, int currentIdToExclude = 0)
		{
			string query = "SELECT COUNT(*) FROM tacgia WHERE HoTen = @HoTen COLLATE NOCASE";
			if (currentIdToExclude > 0)
			{
				query += " AND ID != @ID";
			}
			using (SQLiteCommand cmd = new SQLiteCommand(query, connection))
			{
				cmd.Parameters.AddWithValue("@HoTen", hoTen);
				if (currentIdToExclude > 0)
				{
					cmd.Parameters.AddWithValue("@ID", currentIdToExclude);
				}
				long count = (long)cmd.ExecuteScalar();
				return count == 0;
			}
		}


		private bool IsMaNhaXuatBanUnique(string maNXB, SQLiteConnection connection)
		{
			string checkQuery = "SELECT COUNT(*) FROM NhaXuatBan WHERE MaNhaXuatBan = @MaNhaXuatBan COLLATE NOCASE";
			using (SQLiteCommand checkCmd = new SQLiteCommand(checkQuery, connection))
			{
				checkCmd.Parameters.AddWithValue("@MaNhaXuatBan", maNXB);
				long count = (long)checkCmd.ExecuteScalar();
				return count == 0;
			}
		}
		private bool IsTenNhaXuatBanUnique(string tenNXB, SQLiteConnection connection, int currentIdToExclude = 0)
		{
			string query = "SELECT COUNT(*) FROM NhaXuatBan WHERE TenNhaXuatBan = @TenNhaXuatBan COLLATE NOCASE";
			if (currentIdToExclude > 0)
			{
				query += " AND ID != @ID";
			}
			using (SQLiteCommand cmd = new SQLiteCommand(query, connection))
			{
				cmd.Parameters.AddWithValue("@TenNhaXuatBan", tenNXB);
				if (currentIdToExclude > 0)
				{
					cmd.Parameters.AddWithValue("@ID", currentIdToExclude);
				}
				long count = (long)cmd.ExecuteScalar();
				return count == 0;
			}
		}

		private bool IsTenTheLoaiUnique(string tenTheLoai, SQLiteConnection connection, int currentIdToExclude = 0)
		{
			string query = "SELECT COUNT(*) FROM TheLoai WHERE TenTheLoai = @TenTheLoai COLLATE NOCASE";
			if (currentIdToExclude > 0)
			{
				query += " AND ID != @ID";
			}
			using (SQLiteCommand cmd = new SQLiteCommand(query, connection))
			{
				cmd.Parameters.AddWithValue("@TenTheLoai", tenTheLoai);
				if (currentIdToExclude > 0)
				{
					cmd.Parameters.AddWithValue("@ID", currentIdToExclude);
				}
				long count = (long)cmd.ExecuteScalar();
				return count == 0;
			}
		}


		public bool DoesAuthorExistByName(string hoTen)
		{
			string query = "SELECT COUNT(*) FROM tacgia WHERE HoTen = @HoTen COLLATE NOCASE";
			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						cmd.Parameters.AddWithValue("@HoTen", hoTen ?? (object)DBNull.Value);
						long count = (long)cmd.ExecuteScalar();
						return count > 0;
					}
				}
				catch (Exception ex)
				{
					Console.WriteLine($"Lỗi khi kiểm tra tác giả tồn tại theo tên: {ex.Message}");
					return true;
				}
			}
		}
		public TacGia GetTacGiaByName(string hoTen)
		{
			TacGia author = null;
			string query = "SELECT * FROM tacgia WHERE HoTen = @HoTen COLLATE NOCASE LIMIT 1";
			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						cmd.Parameters.AddWithValue("@HoTen", hoTen);
						using (SQLiteDataReader reader = cmd.ExecuteReader())
						{
							if (reader.Read())
							{
								author = new TacGia
								{
									ID = reader.GetInt32(reader.GetOrdinal("ID")),
									MaTacGia = reader.GetString(reader.GetOrdinal("MaTacGia")),
									HoTen = reader.GetString(reader.GetOrdinal("HoTen")),
									// Lấy các trường khác nếu cần
								};
							}
						}
					}
				}
				catch (Exception ex)
				{
					Console.WriteLine($"Lỗi khi lấy tác giả theo tên '{hoTen}': {ex.Message}");
				}
			}
			return author;
		}


		public bool InsertTacGia(TacGia tacGia, bool generateMaTacGia = true)
		{
			using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
			{
				try
				{
					sqlConnection.Open();

					if (!IsTenTacGiaUnique(tacGia.HoTen, sqlConnection))
					{
						// Không tự động báo lỗi ở đây, để lớp gọi xử lý
						Console.WriteLine($"Tên tác giả '{tacGia.HoTen}' đã tồn tại.");
						return false; // Hoặc có thể throw exception tùy theo thiết kế
					}

					if (generateMaTacGia || string.IsNullOrWhiteSpace(tacGia.MaTacGia))
					{
						string uniqueMaTacGia;
						bool isUnique = false;
						int attempts = 0;
						const int maxAttempts = 20;
						do
						{
							uniqueMaTacGia = GenerateRandomAlphaNumericCode(8, "TG");
							isUnique = IsMaTacGiaUnique(uniqueMaTacGia, sqlConnection);
							attempts++;
							if (attempts > maxAttempts)
							{
								MessageBox.Show($"Không thể tạo Mã Tác Giả duy nhất sau {maxAttempts} lần thử.", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
								return false;
							}
						} while (!isUnique);
						tacGia.MaTacGia = uniqueMaTacGia;
					}
					else // Kiểm tra mã cung cấp có unique không
					{
						if (!IsMaTacGiaUnique(tacGia.MaTacGia, sqlConnection))
						{
							MessageBox.Show($"Mã Tác Giả '{tacGia.MaTacGia}' đã tồn tại.", "Lỗi Trùng Lặp", MessageBoxButtons.OK, MessageBoxIcon.Warning);
							return false;
						}
					}


					string query = @"INSERT INTO tacgia
                            (MaTacGia, HoTen, HocHam, HocVi, DiaChi, Email, NgaySinh, QuocTich, Anh, TieuSu, TrangThai)
                            VALUES
                            (@MaTacGia, @HoTen, @HocHam, @HocVi, @DiaChi, @Email, @NgaySinh, @QuocTich, @Anh, @TieuSu, @TrangThai)";

					using (SQLiteCommand sqlCommand = new SQLiteCommand(query, sqlConnection))
					{
						sqlCommand.Parameters.AddWithValue("@MaTacGia", tacGia.MaTacGia);
						sqlCommand.Parameters.AddWithValue("@HoTen", tacGia.HoTen);
						sqlCommand.Parameters.AddWithValue("@HocHam", (object)tacGia.HocHam ?? DBNull.Value);
						sqlCommand.Parameters.AddWithValue("@HocVi", (object)tacGia.HocVi ?? DBNull.Value);
						sqlCommand.Parameters.AddWithValue("@DiaChi", (object)tacGia.DiaChi ?? DBNull.Value);
						sqlCommand.Parameters.AddWithValue("@Email", (object)tacGia.Email ?? DBNull.Value);
						sqlCommand.Parameters.AddWithValue("@NgaySinh", (object)tacGia.NgaySinh ?? DBNull.Value);
						sqlCommand.Parameters.AddWithValue("@QuocTich", (object)tacGia.QuocTich ?? DBNull.Value);
						sqlCommand.Parameters.AddWithValue("@Anh", (object)tacGia.Anh ?? DBNull.Value);
						sqlCommand.Parameters.AddWithValue("@TieuSu", (object)tacGia.TieuSu ?? DBNull.Value);
						sqlCommand.Parameters.AddWithValue("@TrangThai", (object)tacGia.TrangThai ?? "Hoạt động");

						int rowsAffected = sqlCommand.ExecuteNonQuery();
						return rowsAffected > 0;
					}
				}
				catch (SQLiteException ex)
				{
					if (ex.ErrorCode == (int)SQLiteErrorCode.Constraint_Unique || ex.Message.ToUpper().Contains("UNIQUE CONSTRAINT FAILED"))
					{
						MessageBox.Show($"Lỗi: Mã tác giả '{tacGia.MaTacGia}' hoặc tên tác giả '{tacGia.HoTen}' đã tồn tại.", "Lỗi Trùng Lặp", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					else
					{
						MessageBox.Show($"Lỗi CSDL khi thêm tác giả: {ex.Message} (Mã lỗi: {ex.ErrorCode})", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					Console.WriteLine($"SQLite Error inserting author (Code {ex.ErrorCode}): {ex.Message}");
					return false;
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi không mong muốn khi thêm tác giả: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error inserting author: {ex.Message}");
					return false;
				}
			}
		}

		public bool AddTacGiaIfNotExists(string hoTenTacGia)
		{
			if (string.IsNullOrWhiteSpace(hoTenTacGia))
			{
				return false; // Không thêm nếu tên rỗng
			}

			TacGia existingTacGia = GetTacGiaByName(hoTenTacGia);
			if (existingTacGia == null)
			{
				TacGia newTacGia = new TacGia
				{
					HoTen = hoTenTacGia,
					// Các trường khác có thể để null hoặc giá trị mặc định
					TrangThai = "Hoạt động"
				};
				// Mã tác giả sẽ được tạo tự động trong InsertTacGia
				bool result = InsertTacGia(newTacGia, true);
				if (result)
				{
					Console.WriteLine($"Đã tự động thêm tác giả mới: '{hoTenTacGia}' (Mã: {newTacGia.MaTacGia})");
				}
				return result;
			}
			return true; // Đã tồn tại, coi như thành công
		}


		public List<TacGia> GetAuthorsByName(string hoTen)
		{
			List<TacGia> authors = new List<TacGia>();
			string query = "SELECT ID, MaTacGia, HoTen, NgaySinh, Email, HocVi FROM tacgia WHERE HoTen = @HoTen COLLATE NOCASE ORDER BY MaTacGia";

			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						cmd.Parameters.AddWithValue("@HoTen", hoTen ?? (object)DBNull.Value);
						using (SQLiteDataReader reader = cmd.ExecuteReader())
						{
							while (reader.Read())
							{
								TacGia author = new TacGia();
								author.ID = reader.GetInt32(reader.GetOrdinal("ID"));
								author.MaTacGia = reader.GetString(reader.GetOrdinal("MaTacGia"));
								author.HoTen = reader.GetString(reader.GetOrdinal("HoTen"));
								author.NgaySinh = reader.IsDBNull(reader.GetOrdinal("NgaySinh")) ? null : reader.GetString(reader.GetOrdinal("NgaySinh"));
								author.Email = reader.IsDBNull(reader.GetOrdinal("Email")) ? null : reader.GetString(reader.GetOrdinal("Email"));
								author.HocVi = reader.IsDBNull(reader.GetOrdinal("HocVi")) ? null : reader.GetString(reader.GetOrdinal("HocVi"));
								authors.Add(author);
							}
						}
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi tìm kiếm tác giả theo tên: {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error getting authors by name: {ex.Message}");
				}
			}
			return authors;
		}

		public List<string> GetBookTitlesByAuthorName(string authorName, int limit = 10)
		{
			List<string> bookTitles = new List<string>();
			string query = $"SELECT TenSach FROM quanlysach WHERE TacGia = @AuthorName COLLATE NOCASE ORDER BY TenSach LIMIT {limit}";

			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						cmd.Parameters.AddWithValue("@AuthorName", authorName ?? (object)DBNull.Value);
						using (SQLiteDataReader reader = cmd.ExecuteReader())
						{
							while (reader.Read())
							{
								if (!reader.IsDBNull(0))
								{
									bookTitles.Add(reader.GetString(0));
								}
							}
						}
					}
				}
				catch (Exception ex)
				{
					Console.WriteLine($"Lỗi khi lấy sách theo tên tác giả '{authorName}': {ex.Message}");
				}
			}
			return bookTitles;
		}

		public DataTable GetAllAuthors()
		{
			DataTable dataTable = new DataTable();
			string query = "SELECT ID, MaTacGia, HoTen, HocHam, HocVi, DiaChi, Email, NgaySinh, QuocTich, Anh, TieuSu, TrangThai FROM tacgia ORDER BY HoTen ASC";
			using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
			{
				try
				{
					sqlConnection.Open();
					SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, sqlConnection);
					adapter.Fill(dataTable);
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi lấy danh sách tác giả: {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error getting all authors: {ex.Message}");
				}
			}
			return dataTable;
		}
		public List<string> GetAllAuthorNames()
		{
			List<string> authorNames = new List<string>();
			string query = "SELECT DISTINCT HoTen FROM tacgia WHERE HoTen IS NOT NULL AND HoTen != '' ORDER BY HoTen ASC";
			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						using (SQLiteDataReader reader = cmd.ExecuteReader())
						{
							while (reader.Read())
							{
								authorNames.Add(reader.GetString(0));
							}
						}
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi lấy danh sách tên tác giả: {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error getting all author names: {ex.Message}");
				}
			}
			return authorNames;
		}


		public bool UpdateAuthorImage(int authorId, string relativeImagePath)
		{
			string query = "UPDATE tacgia SET Anh = @ImagePath WHERE ID = @Id";
			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						cmd.Parameters.AddWithValue("@ImagePath", string.IsNullOrEmpty(relativeImagePath) ? (object)DBNull.Value : relativeImagePath);
						cmd.Parameters.AddWithValue("@Id", authorId);
						int rowsAffected = cmd.ExecuteNonQuery();
						return rowsAffected > 0;
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi CSDL khi cập nhật ảnh tác giả (ID: {authorId}): {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"DB Error Updating Author Image (ID: {authorId}): {ex.Message}");
					return false;
				}
			}
		}

		public bool UpdateAuthorBiography(int authorId, string biography)
		{
			string query = "UPDATE tacgia SET TieuSu = @Biography WHERE ID = @Id";
			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						cmd.Parameters.AddWithValue("@Biography", biography ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@Id", authorId);
						int rowsAffected = cmd.ExecuteNonQuery();
						return rowsAffected > 0;
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi CSDL khi cập nhật tiểu sử tác giả (ID: {authorId}): {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"DB Error Updating Author Biography (ID: {authorId}): {ex.Message}");
					return false;
				}
			}
		}

		public bool UpdateAuthorGeneralInfo(TacGia author)
		{
			if (author == null || author.ID <= 0)
			{
				MessageBox.Show("Thông tin tác giả không hợp lệ để cập nhật (Thiếu ID).", "Lỗi Dữ liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return false;
			}

			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					if (!IsTenTacGiaUnique(author.HoTen, conn, author.ID))
					{
						MessageBox.Show($"Tên tác giả '{author.HoTen}' đã được sử dụng bởi tác giả khác.", "Lỗi Trùng Lặp", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return false;
					}

					string query = @"UPDATE tacgia
    						 SET HoTen = @HoTen,
    							 NgaySinh = @NgaySinh,
    							 DiaChi = @DiaChi,
    							 QuocTich = @QuocTich,
    							 Email = @Email,
    							 TrangThai = @TrangThai,
    							 HocHam = @HocHam,
    							 HocVi = @HocVi
    						 WHERE ID = @ID";

					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						cmd.Parameters.AddWithValue("@HoTen", author.HoTen ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@NgaySinh", string.IsNullOrEmpty(author.NgaySinh) ? (object)DBNull.Value : author.NgaySinh);
						cmd.Parameters.AddWithValue("@DiaChi", author.DiaChi ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@QuocTich", author.QuocTich ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@Email", author.Email ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@TrangThai", author.TrangThai ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@HocHam", author.HocHam ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@HocVi", author.HocVi ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@ID", author.ID);

						int rowsAffected = cmd.ExecuteNonQuery();
						return rowsAffected > 0;
					}
				}
				catch (SQLiteException ex)
				{
					MessageBox.Show($"Lỗi CSDL khi cập nhật thông tin tác giả (ID: {author.ID}): {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"DB Error Updating Author Info (ID: {author.ID}): {ex.Message}");
					return false;
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi không mong muốn khi cập nhật thông tin tác giả (ID: {author.ID}): {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error Updating Author Info (ID: {author.ID}): {ex.Message}");
					return false;
				}
			}
		}

		public DataTable GetBooksByAuthorName(string authorName)
		{
			DataTable dataTable = new DataTable();
			string query = @"SELECT ID, Masach, TenSach, TheLoai, TacGia, SoLuong, NhaXuatBan, NamXuatBan
                       FROM quanlysach
                       WHERE TacGia = @AuthorName COLLATE NOCASE
                       ORDER BY TenSach ASC";

			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn);
					adapter.SelectCommand.Parameters.AddWithValue("@AuthorName", authorName ?? (object)DBNull.Value);
					adapter.Fill(dataTable);
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi lấy sách của tác giả '{authorName}': {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error getting books by author name '{authorName}': {ex.Message}");
				}
			}
			return dataTable;
		}

		public DataTable GetBooksByCategoryName(string categoryName)
		{
			DataTable dataTable = new DataTable();
			string query = @"SELECT ID, Masach, TenSach, TheLoai, TacGia, SoLuong, NhaXuatBan, NamXuatBan, anh, pdf, tomtatnoidung
                       FROM quanlysach
                       WHERE TheLoai = @CategoryName COLLATE NOCASE
                       ORDER BY TenSach ASC";

			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn);
					adapter.SelectCommand.Parameters.AddWithValue("@CategoryName", string.IsNullOrEmpty(categoryName) ? (object)DBNull.Value : categoryName);
					adapter.Fill(dataTable);
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi lấy sách theo thể loại '{categoryName}': {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error getting books by category name '{categoryName}': {ex.Message}");
				}
			}
			return dataTable;
		}


		internal string GenerateRandomMaSach(int length = 9)
		{
			const string chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
			return new string(Enumerable.Repeat(chars, length)
			  .Select(s => s[randomGenerator.Next(s.Length)]).ToArray());
		}

		internal bool CheckDuplicateBook(sachquanly sach)
		{
			string query = @"SELECT COUNT(*) FROM quanlysach
                             WHERE TenSach = @TenSach COLLATE NOCASE
                               AND TheLoai = @TheLoai COLLATE NOCASE
                               AND TacGia = @TacGia COLLATE NOCASE
                               AND NhaXuatBan = @NhaXuatBan COLLATE NOCASE
                               AND NamXuatBan = @NamXuatBan";

			using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
			{
				try
				{
					sqlConnection.Open();
					using (SQLiteCommand sqlCommand = new SQLiteCommand(query, sqlConnection))
					{
						sqlCommand.Parameters.AddWithValue("@TenSach", sach.Tensach ?? (object)DBNull.Value);
						sqlCommand.Parameters.AddWithValue("@TheLoai", sach.Theloai ?? (object)DBNull.Value);
						sqlCommand.Parameters.AddWithValue("@TacGia", sach.Tacgia ?? (object)DBNull.Value);
						sqlCommand.Parameters.AddWithValue("@NhaXuatBan", sach.Nhaxuatban ?? (object)DBNull.Value);
						sqlCommand.Parameters.AddWithValue("@NamXuatBan", sach.Namxuatban);

						long count = (long)sqlCommand.ExecuteScalar();
						return count > 0;
					}
				}
				catch (Exception ex)
				{
					Console.WriteLine($"Lỗi khi kiểm tra sách trùng lặp: {ex.Message}");
					MessageBox.Show($"Lỗi khi kiểm tra sách trùng lặp: {ex.Message}", "Lỗi Cơ Sở Dữ Liệu", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return true;
				}
			}
		}

		public DataTable getAllbooks()
		{
			DataTable dataTable = new DataTable();
			try
			{
				string query = "SELECT ID, Masach, TenSach, TheLoai, TacGia, SoLuong, NhaXuatBan, NamXuatBan, anh, pdf, tomtatnoidung FROM quanlysach ORDER BY TenSach ASC";
				using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
				{
					sqlConnection.Open();
					dataAdapter = new SQLiteDataAdapter(query, sqlConnection);
					dataAdapter.Fill(dataTable);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi lấy danh sách sách: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			return dataTable;
		}

		public bool UpdateBookPdfPath(int bookId, string pdfPath)
		{
			string query = "UPDATE quanlysach SET pdf = @PdfPath WHERE ID = @ID";
			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						cmd.Parameters.AddWithValue("@PdfPath", pdfPath ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@ID", bookId);
						int rowsAffected = cmd.ExecuteNonQuery();
						return rowsAffected > 0;
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi cập nhật đường dẫn PDF trong CSDL: {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"DB Error Updating PDF Path: {ex.Message}");
					return false;
				}
			}
		}

		public DataTable GetAllTheLoai()
		{
			DataTable dataTable = new DataTable();
			string query = "SELECT ID, MaTheLoai, TenTheLoai, MoTa, TrangThai FROM TheLoai ORDER BY TenTheLoai ASC";
			using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
			{
				try
				{
					sqlConnection.Open();
					SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, sqlConnection);
					adapter.Fill(dataTable);
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi lấy danh sách thể loại: {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error getting all categories: {ex.Message}");
				}
			}
			return dataTable;
		}
		public List<string> GetAllTenTheLoai()
		{
			List<string> tenTheLoaiList = new List<string>();
			string query = "SELECT DISTINCT TenTheLoai FROM TheLoai WHERE TenTheLoai IS NOT NULL AND TenTheLoai != '' ORDER BY TenTheLoai ASC";
			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						using (SQLiteDataReader reader = cmd.ExecuteReader())
						{
							while (reader.Read())
							{
								tenTheLoaiList.Add(reader.GetString(0));
							}
						}
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi lấy danh sách tên thể loại: {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error getting all category names: {ex.Message}");
				}
			}
			return tenTheLoaiList;
		}


		public bool InsertTheLoai(string maTheLoai, string tenTheLoai, string moTa, string trangThai)
		{
			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					if (!IsTenTheLoaiUnique(tenTheLoai, conn))
					{
						// Không báo lỗi ở đây, để lớp gọi xử lý
						Console.WriteLine($"Tên thể loại '{tenTheLoai}' đã tồn tại.");
						return false;
					}

					string checkQuery = "SELECT COUNT(*) FROM TheLoai WHERE MaTheLoai = @MaTheLoai COLLATE NOCASE";
					using (SQLiteCommand checkCmd = new SQLiteCommand(checkQuery, conn))
					{
						checkCmd.Parameters.AddWithValue("@MaTheLoai", maTheLoai);
						long count = (long)checkCmd.ExecuteScalar();
						if (count > 0)
						{
							MessageBox.Show($"Mã thể loại '{maTheLoai}' đã tồn tại.", "Lỗi Trùng Lặp", MessageBoxButtons.OK, MessageBoxIcon.Warning);
							return false;
						}
					}

					string query = @"INSERT INTO TheLoai (MaTheLoai, TenTheLoai, MoTa, TrangThai)
                                     VALUES (@MaTheLoai, @TenTheLoai, @MoTa, @TrangThai)";
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						cmd.Parameters.AddWithValue("@MaTheLoai", maTheLoai);
						cmd.Parameters.AddWithValue("@TenTheLoai", tenTheLoai ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@MoTa", string.IsNullOrEmpty(moTa) ? (object)DBNull.Value : moTa);
						cmd.Parameters.AddWithValue("@TrangThai", string.IsNullOrEmpty(trangThai) ? (object)DBNull.Value : "Hoạt động");

						int rowsAffected = cmd.ExecuteNonQuery();
						return rowsAffected > 0;
					}
				}
				catch (SQLiteException ex)
				{
					if (ex.ErrorCode == (int)SQLiteErrorCode.Constraint_Unique || ex.Message.ToUpper().Contains("UNIQUE CONSTRAINT FAILED"))
					{
						MessageBox.Show($"Lỗi: Mã thể loại '{maTheLoai}' hoặc tên thể loại '{tenTheLoai}' đã tồn tại.", "Lỗi Trùng Lặp", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					else
					{
						Console.WriteLine($"Lỗi CSDL khi thêm thể loại: {ex.Message} (Mã lỗi: {ex.ErrorCode})");
					}
					return false;
				}
				catch (Exception ex)
				{
					Console.WriteLine($"Lỗi không mong muốn khi thêm thể loại: {ex.Message}");
					throw;
				}
			}
		}

		public bool UpdateTheLoai(int id, string maTheLoai, string tenTheLoai, string moTa, string trangThai)
		{
			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					if (!IsTenTheLoaiUnique(tenTheLoai, conn, id))
					{
						MessageBox.Show($"Tên thể loại '{tenTheLoai}' đã được sử dụng bởi thể loại khác.", "Lỗi Trùng Lặp", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return false;
					}

					string checkDuplicateQuery = "SELECT COUNT(*) FROM TheLoai WHERE MaTheLoai = @MaTheLoai COLLATE NOCASE AND ID != @ID";
					using (SQLiteCommand checkCmd = new SQLiteCommand(checkDuplicateQuery, conn))
					{
						checkCmd.Parameters.AddWithValue("@MaTheLoai", maTheLoai);
						checkCmd.Parameters.AddWithValue("@ID", id);
						long count = (long)checkCmd.ExecuteScalar();
						if (count > 0)
						{
							MessageBox.Show($"Mã thể loại '{maTheLoai}' đã được sử dụng bởi thể loại khác.", "Lỗi Trùng Lặp", MessageBoxButtons.OK, MessageBoxIcon.Error);
							Console.WriteLine($"Lỗi CSDL khi cập nhật thể loại (ID: {id}): Mã thể loại '{maTheLoai}' đã được sử dụng bởi thể loại khác.");
							return false;
						}
					}

					string query = @"UPDATE TheLoai
                             SET MaTheLoai = @MaTheLoai,
                                 TenTheLoai = @TenTheLoai,
                                 MoTa = @MoTa,
                                 TrangThai = @TrangThai
                             WHERE ID = @ID";

					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						cmd.Parameters.AddWithValue("@ID", id);
						cmd.Parameters.AddWithValue("@MaTheLoai", maTheLoai);
						cmd.Parameters.AddWithValue("@TenTheLoai", tenTheLoai ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@MoTa", string.IsNullOrEmpty(moTa) ? (object)DBNull.Value : moTa);
						cmd.Parameters.AddWithValue("@TrangThai", string.IsNullOrEmpty(trangThai) ? (object)DBNull.Value : trangThai);

						int rowsAffected = cmd.ExecuteNonQuery();
						return rowsAffected > 0;
					}
				}
				catch (SQLiteException ex)
				{
					MessageBox.Show($"Lỗi CSDL khi cập nhật thể loại (ID: {id}): {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Lỗi CSDL khi cập nhật thể loại (ID: {id}): {ex.Message} (Mã lỗi: {ex.ErrorCode})");
					return false;
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi không mong muốn khi cập nhật thể loại (ID: {id}): {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Lỗi không mong muốn khi cập nhật thể loại (ID: {id}): {ex.Message}");
					throw;
				}
			}
		}


		public bool UpdateBookImagePath(int bookId, string imagePath)
		{
			string query = "UPDATE quanlysach SET anh = @ImagePath WHERE ID = @ID";
			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						cmd.Parameters.AddWithValue("@ImagePath", imagePath ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@ID", bookId);
						int rowsAffected = cmd.ExecuteNonQuery();
						return rowsAffected > 0;
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi cập nhật đường dẫn ảnh trong CSDL: {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"DB Error Updating Image Path: {ex.Message}");
					return false;
				}
			}
		}

		public string GetBookPdfPath(int bookId)
		{
			string query = "SELECT pdf FROM quanlysach WHERE ID = @ID";
			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						cmd.Parameters.AddWithValue("@ID", bookId);
						object result = cmd.ExecuteScalar();
						if (result != null && result != DBNull.Value)
						{
							return result.ToString();
						}
						else
						{
							return null;
						}
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi lấy đường dẫn PDF từ CSDL: {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"DB Error Getting PDF Path: {ex.Message}");
					return null;
				}
			}
		}

		public bool Insert(sachquanly sachquanly)
		{
			if (CheckDuplicateBook(sachquanly))
			{
				if (!Application.OpenForms.OfType<Form>().Any(f => f is Form && f.Text == "Lỗi Cơ Sở Dữ Liệu"))
				{
					MessageBox.Show("Sách này (dựa trên Tên, Thể loại, Tác giả, NXB, Năm XB) đã tồn tại.", "Lỗi Trùng Lặp", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
				return false;
			}

			if (string.IsNullOrWhiteSpace(sachquanly.Masach))
			{
				string newMaSach;
				bool isMaSachUnique = false;
				using (SQLiteConnection checkConnection = Connection.GetSQLiteConnection())
				{
					try
					{
						checkConnection.Open();
						int attempts = 0;
						const int maxAttempts = 20;
						do
						{
							newMaSach = GenerateRandomMaSach();
							string checkQuery = "SELECT COUNT(*) FROM quanlysach WHERE Masach = @Masach";
							using (SQLiteCommand checkCmd = new SQLiteCommand(checkQuery, checkConnection))
							{
								checkCmd.Parameters.AddWithValue("@Masach", newMaSach);
								long count = (long)checkCmd.ExecuteScalar();
								if (count == 0)
								{
									isMaSachUnique = true;
								}
							}
							attempts++;
							if (attempts > maxAttempts)
							{
								throw new Exception($"Không thể tạo mã sách duy nhất sau {maxAttempts} lần thử.");
							}
						} while (!isMaSachUnique);
						sachquanly.Masach = newMaSach;
					}
					catch (Exception ex)
					{
						MessageBox.Show($"Lỗi khi tạo mã sách duy nhất: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
						return false;
					}
				}
			}

			string query = @"INSERT INTO quanlysach
                    (Masach, TenSach, TheLoai, TacGia, SoLuong, NhaXuatBan, NamXuatBan)
                    VALUES
                    (@Masach, @TenSach, @TheLoai, @TacGia, @SoLuong, @NhaXuatBan, @NamXuatBan)";

			using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
			{
				try
				{
					sqlConnection.Open();
					using (SQLiteCommand sqlCommand = new SQLiteCommand(query, sqlConnection))
					{
						sqlCommand.Parameters.AddWithValue("@Masach", sachquanly.Masach);
						sqlCommand.Parameters.AddWithValue("@TenSach", sachquanly.Tensach);
						sqlCommand.Parameters.AddWithValue("@TheLoai", sachquanly.Theloai);
						sqlCommand.Parameters.AddWithValue("@TacGia", sachquanly.Tacgia);
						sqlCommand.Parameters.AddWithValue("@SoLuong", sachquanly.Soluong);
						sqlCommand.Parameters.AddWithValue("@NhaXuatBan", sachquanly.Nhaxuatban);
						sqlCommand.Parameters.AddWithValue("@NamXuatBan", sachquanly.Namxuatban);

						int rowsAffected = sqlCommand.ExecuteNonQuery();
						return rowsAffected > 0;
					}
				}
				catch (SQLiteException ex)
				{
					if (ex.ErrorCode == (int)SQLiteErrorCode.Constraint_Unique || ex.Message.Contains("UNIQUE constraint failed"))
					{
						MessageBox.Show($"Lỗi: Mã sách '{sachquanly.Masach}' đã tồn tại hoặc có lỗi trùng lặp khác. Vui lòng kiểm tra lại.", "Lỗi Trùng Lặp", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					else
					{
						MessageBox.Show($"Lỗi cơ sở dữ liệu khi thêm sách: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					Console.WriteLine($"Lỗi SQLite khi thêm sách: {ex.Message}");
					return false;
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi không xác định khi thêm sách: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Lỗi không xác định khi thêm sách: {ex.Message}");
					return false;
				}
			}
		}

		public bool Update(sachquanly sachquanly)
		{
			string query = @"UPDATE quanlysach
                     SET TenSach = @TenSach,
                         TheLoai = @TheLoai,
                         TacGia = @TacGia,
                         SoLuong = @SoLuong,
                         NhaXuatBan = @NhaXuatBan,
                         NamXuatBan = @NamXuatBan
                     WHERE ID = @ID";

			using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
			{
				try
				{
					sqlConnection.Open();
					using (SQLiteCommand sqlCommand = new SQLiteCommand(query, sqlConnection))
					{
						sqlCommand.Parameters.AddWithValue("@ID", sachquanly.ID);
						sqlCommand.Parameters.AddWithValue("@TenSach", sachquanly.Tensach);
						sqlCommand.Parameters.AddWithValue("@TheLoai", sachquanly.Theloai);
						sqlCommand.Parameters.AddWithValue("@TacGia", sachquanly.Tacgia);
						sqlCommand.Parameters.AddWithValue("@SoLuong", sachquanly.Soluong);
						sqlCommand.Parameters.AddWithValue("@NhaXuatBan", sachquanly.Nhaxuatban);
						sqlCommand.Parameters.AddWithValue("@NamXuatBan", sachquanly.Namxuatban);

						int rowsAffected = sqlCommand.ExecuteNonQuery();
						return rowsAffected > 0;
					}
				}
				catch (Exception ex)
				{
					Console.WriteLine($"Lỗi khi cập nhật sách: {ex.Message}");
					MessageBox.Show($"Lỗi khi cập nhật sách: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return false;
				}
			}
		}
		public bool delete(int ID)
		{
			SQLiteConnection sqlConnection = Connection.GetSQLiteConnection();
			string query = "DELETE FROM quanlysach WHERE ID = @ID";

			try
			{
				sqlConnection.Open();
				sqlcommand = new SQLiteCommand(query, sqlConnection);
				sqlcommand.Parameters.AddWithValue("@ID", ID);
				sqlcommand.ExecuteNonQuery();
				return true;
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi xóa sách (ID: {ID}): {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
			finally
			{
				if (sqlConnection.State == ConnectionState.Open)
				{
					sqlConnection.Close();
				}
			}
		}

		public List<sachquanly> ImportFromExcel(string filePath)
		{
			var books = new List<sachquanly>();

			System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
			using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
			{
				using (var reader = ExcelReaderFactory.CreateReader(stream))
				{
					var result = reader.AsDataSet();
					if (result.Tables.Count > 0)
					{
						var table = result.Tables[0];
						for (int i = 1; i < table.Rows.Count; i++)
						{
							var row = table.Rows[i];
							try
							{
								var book = new sachquanly
								{
									Masach = row[0].ToString(),
									Tensach = row[1].ToString(),
									Theloai = row[2].ToString(),
									Tacgia = row[3].ToString(),
									Soluong = Convert.ToInt32(row[4]),
									Nhaxuatban = row[5].ToString(),
									Namxuatban = Convert.ToInt32(row[6])
								};
								books.Add(book);
							}
							catch (Exception ex)
							{
								Console.WriteLine($"Lỗi khi đọc dòng {i + 1} từ Excel: {ex.Message}. Dữ liệu: {string.Join(", ", row.ItemArray)}");
							}
						}
					}
				}
			}
			return books;
		}
		public DataTable searchBooks(string keyword, SQLiteConnection sqlConnection = null)
		{
			DataTable dt = new DataTable();
			string query = @"SELECT ID, Masach, TenSach, TheLoai, TacGia, SoLuong, NhaXuatBan, NamXuatBan, anh, pdf, tomtatnoidung
                           FROM quanlysach
                           WHERE TenSach LIKE @Keyword COLLATE NOCASE
                              OR Masach LIKE @Keyword COLLATE NOCASE
                              OR TacGia LIKE @Keyword COLLATE NOCASE 
                              OR TheLoai LIKE @Keyword COLLATE NOCASE
                              OR NhaXuatBan LIKE @Keyword COLLATE NOCASE";
			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					dataAdapter = new SQLiteDataAdapter(query, conn);
					dataAdapter.SelectCommand.Parameters.AddWithValue("@Keyword", "%" + keyword + "%");
					dataAdapter.Fill(dt);
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi tìm kiếm sách: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
			}
			return dt;
		}
		public DataTable getAllphieu()
		{
			DataTable dataTable = new DataTable();
			string query = "SELECT * FROM muontrasach ORDER BY TenSach ASC";
			using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
			{
				sqlConnection.Open();
				dataAdapter = new SQLiteDataAdapter(query, sqlConnection);
				dataAdapter.Fill(dataTable);
				sqlConnection.Close();
			}
			return dataTable;
		}

		public bool themphieu(muontra muontra)
		{
			string query = @"INSERT INTO muontrasach (TenSach, MaSinhVien, HoVaTen, LoaiPhieu, NgayMuon, NgayTra, SoLuong, MaSach) 
                    VALUES (@TenSach, @MaSinhVien, @HoVaTen, @LoaiPhieu, @NgayMuon, @NgayTra, @SoLuong, @MaSach)";

			using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
			{
				try
				{
					sqlConnection.Open();
					using (SQLiteCommand sqlCommand = new SQLiteCommand(query, sqlConnection))
					{
						sqlCommand.Parameters.Add("@TenSach", DbType.String).Value = muontra.Tensach;
						sqlCommand.Parameters.Add("@MaSinhVien", DbType.String).Value = muontra.Masinhvien;
						sqlCommand.Parameters.Add("@HoVaTen", DbType.String).Value = muontra.Hovaten;
						sqlCommand.Parameters.Add("@LoaiPhieu", DbType.String).Value = muontra.Loaiphieu;
						sqlCommand.Parameters.Add("@NgayMuon", DbType.Date).Value = muontra.Ngaymuon;
						sqlCommand.Parameters.Add("@NgayTra", DbType.Date).Value = muontra.Ngaytra;
						sqlCommand.Parameters.Add("@SoLuong", DbType.Int32).Value = muontra.Soluong;
						sqlCommand.Parameters.Add("@MaSach", DbType.String).Value = muontra.Masach;

						int rowsAffected = sqlCommand.ExecuteNonQuery();
						return rowsAffected > 0;
					}
				}
				catch (Exception ex)
				{
					Console.WriteLine($"Lỗi: {ex.Message}");
					MessageBox.Show($"Lỗi khi thêm phiếu mượn: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return false;
				}
			}
		}


		public bool suaphieu(muontra muontra)
		{
			string query = @"UPDATE muontrasach SET TenSach = @TenSach, MaSinhVien = @MaSinhVien, HoVaTen = @HoVaTen, 
                    LoaiPhieu = @LoaiPhieu, NgayMuon = @NgayMuon, NgayTra = @NgayTra, SoLuong = @SoLuong, MaSach = @MaSach 
                    WHERE ID = @ID";

			using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
			{
				try
				{
					sqlConnection.Open();
					using (SQLiteCommand sqlCommand = new SQLiteCommand(query, sqlConnection))
					{
						sqlCommand.Parameters.AddWithValue("@ID", muontra.ID);
						sqlCommand.Parameters.Add("@TenSach", DbType.String).Value = muontra.Tensach;
						sqlCommand.Parameters.Add("@MaSinhVien", DbType.String).Value = muontra.Masinhvien;
						sqlCommand.Parameters.Add("@HoVaTen", DbType.String).Value = muontra.Hovaten;
						sqlCommand.Parameters.Add("@LoaiPhieu", DbType.String).Value = muontra.Loaiphieu;
						sqlCommand.Parameters.Add("@NgayMuon", DbType.Date).Value = muontra.Ngaymuon;
						sqlCommand.Parameters.Add("@NgayTra", DbType.Date).Value = muontra.Ngaytra;
						sqlCommand.Parameters.Add("@SoLuong", DbType.Int32).Value = muontra.Soluong;
						sqlCommand.Parameters.Add("@MaSach", DbType.String).Value = muontra.Masach;


						int rowsAffected = sqlCommand.ExecuteNonQuery();
						return rowsAffected > 0;
					}
				}
				catch (Exception ex)
				{
					Console.WriteLine($"Lỗi: {ex.Message}");
					MessageBox.Show($"Lỗi khi sửa phiếu mượn: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return false;
				}
			}
		}

		public bool trasach(int ID)
		{
			SQLiteConnection sqlConnection = Connection.GetSQLiteConnection();
			string query = "UPDATE muontrasach SET LoaiPhieu = 'Đã trả', NgayTra = @NgayTra, IsActive = 0 WHERE ID = @ID";

			try
			{
				sqlConnection.Open();
				SQLiteCommand sqlcommand = new SQLiteCommand(query, sqlConnection);
				sqlcommand.Parameters.AddWithValue("@ID", ID);
				sqlcommand.Parameters.AddWithValue("@NgayTra", DateTime.Now.ToString("yyyy-MM-dd"));

				sqlcommand.ExecuteNonQuery();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi cập nhật trạng thái trả sách: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
			finally
			{
				if (sqlConnection.State == ConnectionState.Open)
				{
					sqlConnection.Close();
				}
			}
			return true;
		}

		public DataTable timsv(string MaSinhVien, SQLiteConnection sqlConnection = null)
		{
			DataTable dt = new DataTable();
			string query = "select * from muontrasach where MaSinhVien LIKE @MaSinhVien";
			using (SQLiteConnection sqliteConnection = Connection.GetSQLiteConnection())
			{
				try
				{
					sqliteConnection.Open();
					dataAdapter = new SQLiteDataAdapter(query, sqliteConnection);
					dataAdapter.SelectCommand.Parameters.Add("@MaSinhVien", DbType.String).Value = "%" + MaSinhVien + "%";
					dataAdapter.Fill(dt);
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi tìm sinh viên: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				finally
				{
					if (sqliteConnection.State == ConnectionState.Open)
					{
						sqliteConnection.Close();
					}
				}
			}
			return dt;
		}
		public List<TaiKhoan> taiKhoans(string query, Dictionary<string, object> parameters = null)
		{
			List<TaiKhoan> taiKhoans = new List<TaiKhoan>();

			using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
			{
				sqlConnection.Open();
				SQLiteCommand sqlcommand = new SQLiteCommand(query, sqlConnection);

				if (parameters != null)
				{
					foreach (var param in parameters)
					{
						sqlcommand.Parameters.AddWithValue(param.Key, param.Value);
					}
				}

				SQLiteDataReader sqlDataReader = sqlcommand.ExecuteReader();

				while (sqlDataReader.Read())
				{
					TaiKhoan taiKhoan = new TaiKhoan();
					taiKhoan.MaSinhVien = sqlDataReader["MaSinhVien"] != DBNull.Value ? sqlDataReader["MaSinhVien"].ToString() : "";
					taiKhoan.TenTaiKhoan = sqlDataReader["TenTaiKhoan"] != DBNull.Value ? sqlDataReader["TenTaiKhoan"].ToString() : "";
					taiKhoan.MatKhau = sqlDataReader["MatKhau"] != DBNull.Value ? sqlDataReader["MatKhau"].ToString() : "";
					taiKhoan.Email = sqlDataReader["Email"] != DBNull.Value ? sqlDataReader["Email"].ToString() : "";
					taiKhoan.HoVaTen = sqlDataReader["HoVaTen"] != DBNull.Value ? sqlDataReader["HoVaTen"].ToString() : "";
					taiKhoan.Avatar = sqlDataReader["Avatar"] != DBNull.Value ? sqlDataReader["Avatar"].ToString() : "default.png";
					taiKhoan.Quyen = sqlDataReader["Quyen"] != DBNull.Value ? sqlDataReader["Quyen"].ToString() : "user";
					taiKhoan.DiaChi = sqlDataReader["DiaChi"] != DBNull.Value ? sqlDataReader["DiaChi"].ToString() : "";
					try { taiKhoan.SoDienThoai = sqlDataReader["SoDienThoai"] != DBNull.Value ? sqlDataReader["SoDienThoai"].ToString() : ""; }
					catch { taiKhoan.SoDienThoai = ""; }
					try { taiKhoan.NgaySinh = sqlDataReader["NgaySinh"] != DBNull.Value ? sqlDataReader["NgaySinh"].ToString() : ""; }
					catch { taiKhoan.NgaySinh = ""; }
					try { taiKhoan.GioiTinh = sqlDataReader["GioiTinh"] != DBNull.Value ? sqlDataReader["GioiTinh"].ToString() : ""; }
					catch { taiKhoan.GioiTinh = ""; }
					taiKhoans.Add(taiKhoan);
				}
				sqlConnection.Close();
			}
			return taiKhoans;
		}

		public bool Command(string query, Dictionary<string, object> parameters = null)
		{
			try
			{
				using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
				{
					sqlConnection.Open();
					SQLiteCommand sqlcommand = new SQLiteCommand(query, sqlConnection);

					if (parameters != null)
					{
						foreach (var param in parameters)
						{
							sqlcommand.Parameters.AddWithValue(param.Key, param.Value);
						}
					}
					int result = sqlcommand.ExecuteNonQuery();
					sqlConnection.Close();
					return result > 0;
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine("Lỗi thực thi lệnh SQL: " + ex.Message);
				MessageBox.Show("Lỗi thực thi lệnh SQL: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
		}

		public void ExportToExcel(string filePath)
		{
			DataTable dtBooks = getAllbooks();
			DataTable dtBorrowBooks = getAllphieu();

			using (var workbook = new XLWorkbook())
			{
				var worksheetBooks = workbook.Worksheets.Add(dtBooks, "QuanLySach");
				var worksheetBorrowBooks = workbook.Worksheets.Add(dtBorrowBooks, "MuonTraSach");
				workbook.SaveAs(filePath);
			}
		}

		public void ExportToExcelquanlysach(string filePath)
		{
			DataTable dtBooks = getAllbooks();
			using (var workbook = new XLWorkbook())
			{
				var worksheetBooks = workbook.Worksheets.Add(dtBooks, "QuanLySach");
				workbook.SaveAs(filePath);
			}
		}

		public void ExportToExcelmuontrasach(string filePath)
		{
			DataTable dtBorrowBooks = getAllphieu();
			using (var workbook = new XLWorkbook())
			{
				var worksheetBorrowBooks = workbook.Worksheets.Add(dtBorrowBooks, "MuonTraSach");
				workbook.SaveAs(filePath);
			}
		}

		public bool datlichhen(henmuonsach henmuonsach)
		{
			string query = @"INSERT INTO lichhenmuonsach 
                    (MaSV, HoVaTen, TenSach, LichHen, GioHen, LoaiHen) 
                    VALUES 
                    (@MaSV, @HoVaTen, @TenSach, @LichHen, @GioHen, @LoaiHen)";

			using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
			{
				try
				{
					sqlConnection.Open();
					using (SQLiteCommand sqlCommand = new SQLiteCommand(query, sqlConnection))
					{
						sqlCommand.Parameters.Add("@MaSV", DbType.String).Value = henmuonsach.Masinhvien;
						sqlCommand.Parameters.Add("@HoVaTen", DbType.String).Value = henmuonsach.Hovaten;
						sqlCommand.Parameters.Add("@TenSach", DbType.String).Value = henmuonsach.Tensach;
						sqlCommand.Parameters.Add("@LichHen", DbType.String).Value = henmuonsach.Lichhen;
						sqlCommand.Parameters.Add("@GioHen", DbType.String).Value = henmuonsach.Giohen;
						sqlCommand.Parameters.Add("@LoaiHen", DbType.String).Value = henmuonsach.Loaihen;
						int rowsAffected = sqlCommand.ExecuteNonQuery();
						return rowsAffected > 0;
					}
				}
				catch (Exception ex)
				{
					Console.WriteLine($"Lỗi: {ex.Message}");
					MessageBox.Show($"Lỗi khi đặt lịch hẹn: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return false;
				}
			}
		}
		public DataTable getAllhen()
		{
			DataTable dataTable = new DataTable();
			string query = "SELECT * FROM lichhenmuonsach ORDER BY HoVaTen ASC";
			using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
			{
				sqlConnection.Open();
				dataAdapter = new SQLiteDataAdapter(query, sqlConnection);
				dataAdapter.Fill(dataTable);
				sqlConnection.Close();
			}
			return dataTable;
		}

		public DataTable timsvv(string MaSinhVien, SQLiteConnection sqlConnection = null)
		{
			DataTable dt = new DataTable();
			string query = "select * from lichhenmuonsach where MaSV LIKE @MaSV";
			using (SQLiteConnection sqliteConnection = Connection.GetSQLiteConnection())
			{
				try
				{
					sqliteConnection.Open();
					dataAdapter = new SQLiteDataAdapter(query, sqliteConnection);
					dataAdapter.SelectCommand.Parameters.Add("@MaSV", DbType.String).Value = "%" + MaSinhVien + "%";
					dataAdapter.Fill(dt);
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi tìm sinh viên trong lịch hẹn: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				finally
				{
					if (sqliteConnection.State == ConnectionState.Open)
					{
						sqliteConnection.Close();
					}
				}
			}
			return dt;
		}

		public bool dunghen(int ID)
		{
			SQLiteConnection sqlConnection = Connection.GetSQLiteConnection();
			string query = "DELETE FROM lichhenmuonsach WHERE ID = @ID";

			try
			{
				sqlConnection.Open();
				sqlcommand = new SQLiteCommand(query, sqlConnection);
				sqlcommand.Parameters.AddWithValue("@ID", ID);
				sqlcommand.ExecuteNonQuery();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi dừng hẹn: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
				return false;
			}
			finally
			{
				if (sqlConnection.State == ConnectionState.Open)
				{
					sqlConnection.Close();
				}
			}
			return true;
		}

		public void CapNhatThongKeMuon(string maSinhVien, string hoVaTen, int soLuong, DateTime ngayMuon = default)
		{
			try
			{
				if (ngayMuon == default)
					ngayMuon = DateTime.Now;

				using (SQLiteConnection connection = Connection.GetSQLiteConnection())
				{
					connection.Open();
					string checkQuery = "SELECT COUNT(*) FROM ThongKeMuonTra WHERE MaSinhVien = @MaSinhVien";
					SQLiteCommand checkCommand = new SQLiteCommand(checkQuery, connection);
					checkCommand.Parameters.AddWithValue("@MaSinhVien", maSinhVien);
					int count = Convert.ToInt32(checkCommand.ExecuteScalar());

					if (count > 0)
					{
						string updateQuery = @"UPDATE ThongKeMuonTra 
                                    SET SoLanMuon = SoLanMuon + 1, 
                                        TongSachMuon = TongSachMuon + @SoLuong,
                                        NgayMuon = @NgayMuon
                                    WHERE MaSinhVien = @MaSinhVien";
						SQLiteCommand updateCommand = new SQLiteCommand(updateQuery, connection);
						updateCommand.Parameters.AddWithValue("@MaSinhVien", maSinhVien);
						updateCommand.Parameters.AddWithValue("@SoLuong", soLuong);
						updateCommand.Parameters.AddWithValue("@NgayMuon", ngayMuon.ToString("yyyy-MM-dd"));
						updateCommand.ExecuteNonQuery();
					}
					else
					{
						string insertQuery = @"INSERT INTO ThongKeMuonTra 
                                    (MaSinhVien, HoVaTen, SoLanMuon, SoLanTra, TongSachMuon, TongSachTra, NgayMuon) 
                                    VALUES (@MaSinhVien, @HoVaTen, 1, 0, @SoLuong, 0, @NgayMuon)";
						SQLiteCommand insertCommand = new SQLiteCommand(insertQuery, connection);
						insertCommand.Parameters.AddWithValue("@MaSinhVien", maSinhVien);
						insertCommand.Parameters.AddWithValue("@HoVaTen", hoVaTen);
						insertCommand.Parameters.AddWithValue("@SoLuong", soLuong);
						insertCommand.Parameters.AddWithValue("@NgayMuon", ngayMuon.ToString("yyyy-MM-dd"));
						insertCommand.ExecuteNonQuery();
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine("Lỗi khi cập nhật thống kê mượn: " + ex.Message);
				MessageBox.Show("Lỗi khi cập nhật thống kê mượn: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		public void CapNhatThongKeTra(string maSinhVien, string hoVaTen, int soLuong, DateTime ngayTra = default)
		{
			try
			{
				if (ngayTra == default)
					ngayTra = DateTime.Now;

				using (SQLiteConnection connection = Connection.GetSQLiteConnection())
				{
					connection.Open();
					string checkQuery = "SELECT COUNT(*) FROM ThongKeMuonTra WHERE MaSinhVien = @MaSinhVien";
					SQLiteCommand checkCommand = new SQLiteCommand(checkQuery, connection);
					checkCommand.Parameters.AddWithValue("@MaSinhVien", maSinhVien);
					int count = Convert.ToInt32(checkCommand.ExecuteScalar());

					if (count > 0)
					{
						string updateQuery = @"UPDATE ThongKeMuonTra 
                                    SET SoLanTra = SoLanTra + 1, 
                                        TongSachTra = TongSachTra + @SoLuong,
                                        NgayTra = @NgayTra
                                    WHERE MaSinhVien = @MaSinhVien";
						SQLiteCommand updateCommand = new SQLiteCommand(updateQuery, connection);
						updateCommand.Parameters.AddWithValue("@MaSinhVien", maSinhVien);
						updateCommand.Parameters.AddWithValue("@SoLuong", soLuong);
						updateCommand.Parameters.AddWithValue("@NgayTra", ngayTra.ToString("yyyy-MM-dd"));
						updateCommand.ExecuteNonQuery();
					}
					else
					{
						string insertQuery = @"INSERT INTO ThongKeMuonTra 
                                    (MaSinhVien, HoVaTen, SoLanMuon, SoLanTra, TongSachMuon, TongSachTra, NgayTra) 
                                    VALUES (@MaSinhVien, @HoVaTen, 0, 1, 0, @SoLuong, @NgayTra)";
						SQLiteCommand insertCommand = new SQLiteCommand(insertQuery, connection);
						insertCommand.Parameters.AddWithValue("@MaSinhVien", maSinhVien);
						insertCommand.Parameters.AddWithValue("@HoVaTen", hoVaTen);
						insertCommand.Parameters.AddWithValue("@SoLuong", soLuong);
						insertCommand.Parameters.AddWithValue("@NgayTra", ngayTra.ToString("yyyy-MM-dd"));
						insertCommand.ExecuteNonQuery();
					}
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine("Lỗi khi cập nhật thống kê trả: " + ex.Message);
				MessageBox.Show("Lỗi khi cập nhật thống kê trả: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		public DataTable GetThongKeMuonTra(string loaiThongKe, string datePattern, bool isDetailView = false)
		{
			DataTable dt = new DataTable();
			try
			{
				using (SQLiteConnection connection = Connection.GetSQLiteConnection())
				{
					connection.Open();
					string query;
					string dateColumn = loaiThongKe == "Thống kê mượn sách" ? "NgayMuon" : "NgayTra";

					if (isDetailView)
					{
						if (loaiThongKe == "Thống kê mượn sách")
						{
							query = $@"SELECT MaSinhVien, HoVaTen, SoLanMuon, TongSachMuon, {dateColumn} 
                            FROM ThongKeMuonTra 
                            WHERE {dateColumn} LIKE @DatePattern OR @DatePattern = '%'
                            ORDER BY SoLanMuon DESC, TongSachMuon DESC";
						}
						else
						{
							query = $@"SELECT MaSinhVien, HoVaTen, SoLanTra, TongSachTra, {dateColumn} 
                            FROM ThongKeMuonTra 
                            WHERE {dateColumn} LIKE @DatePattern OR @DatePattern = '%'
                            ORDER BY SoLanTra DESC, TongSachTra DESC";
						}
					}
					else
					{
						if (loaiThongKe == "Thống kê mượn sách")
						{
							query = $@"SELECT MaSinhVien, HoVaTen, SoLanMuon, TongSachMuon 
                            FROM ThongKeMuonTra 
                            WHERE {dateColumn} LIKE @DatePattern OR @DatePattern = '%'
                            ORDER BY SoLanMuon DESC, TongSachMuon DESC";
						}
						else
						{
							query = $@"SELECT MaSinhVien, HoVaTen, SoLanTra, TongSachTra 
                            FROM ThongKeMuonTra 
                            WHERE {dateColumn} LIKE @DatePattern OR @DatePattern = '%'
                            ORDER BY SoLanTra DESC, TongSachTra DESC";
						}
					}
					SQLiteCommand command = new SQLiteCommand(query, connection);
					command.Parameters.AddWithValue("@DatePattern", datePattern);
					SQLiteDataAdapter adapter = new SQLiteDataAdapter(command);
					adapter.Fill(dt);
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine("Lỗi khi lấy thống kê: " + ex.Message);
				MessageBox.Show("Lỗi khi lấy thống kê: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			return dt;
		}

		public DataTable TimKiemSinhVienThongKe(string maSinhVien, string loaiThongKe, bool isDetailView = false)
		{
			DataTable dt = new DataTable();
			try
			{
				using (SQLiteConnection connection = Connection.GetSQLiteConnection())
				{
					connection.Open();
					string query;
					string dateColumn = loaiThongKe == "Thống kê mượn sách" ? "NgayMuon" : "NgayTra";

					if (isDetailView)
					{
						if (loaiThongKe == "Thống kê mượn sách")
						{
							query = $@"SELECT MaSinhVien, HoVaTen, SoLanMuon, TongSachMuon, {dateColumn} 
                            FROM ThongKeMuonTra 
                            WHERE MaSinhVien LIKE @MaSinhVien 
                            ORDER BY SoLanMuon DESC, TongSachMuon DESC";
						}
						else
						{
							query = $@"SELECT MaSinhVien, HoVaTen, SoLanTra, TongSachTra, {dateColumn} 
                            FROM ThongKeMuonTra 
                            WHERE MaSinhVien LIKE @MaSinhVien 
                            ORDER BY SoLanTra DESC, TongSachTra DESC";
						}
					}
					else
					{
						if (loaiThongKe == "Thống kê mượn sách")
						{
							query = @"SELECT MaSinhVien, HoVaTen, SoLanMuon, TongSachMuon 
                            FROM ThongKeMuonTra 
                            WHERE MaSinhVien LIKE @MaSinhVien 
                            ORDER BY SoLanMuon DESC, TongSachMuon DESC";
						}
						else
						{
							query = @"SELECT MaSinhVien, HoVaTen, SoLanTra, TongSachTra 
                            FROM ThongKeMuonTra 
                            WHERE MaSinhVien LIKE @MaSinhVien 
                            ORDER BY SoLanTra DESC, TongSachTra DESC";
						}
					}
					SQLiteCommand command = new SQLiteCommand(query, connection);
					command.Parameters.AddWithValue("@MaSinhVien", "%" + maSinhVien + "%");
					SQLiteDataAdapter adapter = new SQLiteDataAdapter(command);
					adapter.Fill(dt);
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine("Lỗi khi tìm kiếm sinh viên: " + ex.Message);
				MessageBox.Show("Lỗi khi tìm kiếm sinh viên trong thống kê: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			return dt;
		}

		public DataTable GetAllThongKeMuonTra(bool isDetailView = false)
		{
			DataTable dt = new DataTable();
			try
			{
				using (SQLiteConnection connection = Connection.GetSQLiteConnection())
				{
					connection.Open();
					string query;
					if (isDetailView)
					{
						query = @"SELECT MaSinhVien, HoVaTen, SoLanMuon, TongSachMuon, 
                         SoLanTra, TongSachTra, NgayMuon, NgayTra 
                         FROM ThongKeMuonTra 
                         ORDER BY SoLanMuon DESC, TongSachMuon DESC";
					}
					else
					{
						query = @"SELECT MaSinhVien, HoVaTen, SoLanMuon, TongSachMuon, 
                         SoLanTra, TongSachTra 
                         FROM ThongKeMuonTra 
                         ORDER BY SoLanMuon DESC, TongSachMuon DESC";
					}
					SQLiteCommand command = new SQLiteCommand(query, connection);
					SQLiteDataAdapter adapter = new SQLiteDataAdapter(command);
					adapter.Fill(dt);
				}
			}
			catch (Exception ex)
			{
				Console.WriteLine("Lỗi khi lấy tất cả thống kê: " + ex.Message);
				MessageBox.Show("Lỗi khi lấy tất cả thống kê: " + ex.Message, "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			return dt;
		}

		public List<string> GetDistinctValues(string columnName)
		{
			List<string> values = new List<string>();
			try
			{
				string query = $"SELECT DISTINCT {columnName} FROM quanlysach WHERE {columnName} IS NOT NULL AND {columnName} != '' ORDER BY {columnName}";
				using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
				{
					sqlConnection.Open();
					SQLiteCommand command = new SQLiteCommand(query, sqlConnection);
					SQLiteDataReader reader = command.ExecuteReader();
					while (reader.Read())
					{
						if (!reader.IsDBNull(0))
							values.Add(reader[0].ToString());
					}
					sqlConnection.Close();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi lấy dữ liệu distinct: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			return values;
		}

		public DataTable SearchBooksByColumn(string columnName, string value)
		{
			DataTable dataTable = new DataTable();
			try
			{
				string query = $"SELECT ID, Masach, TenSach, TheLoai, TacGia, SoLuong, NhaXuatBan, NamXuatBan, anh, pdf, tomtatnoidung FROM quanlysach WHERE {columnName} = @Value ORDER BY TenSach ASC";
				using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
				{
					sqlConnection.Open();
					dataAdapter = new SQLiteDataAdapter(query, sqlConnection);
					dataAdapter.SelectCommand.Parameters.Add("@Value", DbType.String).Value = value;
					dataAdapter.Fill(dataTable);
					sqlConnection.Close();
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi tìm kiếm sách theo cột: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			return dataTable;
		}

		public bool AddTheLoaiIfNotExists(string tenTheLoaiInput)
		{
			if (string.IsNullOrWhiteSpace(tenTheLoaiInput))
			{
				return true;
			}

			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					string checkTenQuery = "SELECT COUNT(*) FROM TheLoai WHERE TenTheLoai = @TenTheLoai COLLATE NOCASE";
					using (SQLiteCommand checkTenCmd = new SQLiteCommand(checkTenQuery, conn))
					{
						checkTenCmd.Parameters.AddWithValue("@TenTheLoai", tenTheLoaiInput);
						long tenCount = (long)checkTenCmd.ExecuteScalar();
						if (tenCount > 0)
						{
							Console.WriteLine($"Thể loại '{tenTheLoaiInput}' đã tồn tại trong bảng TheLoai.");
							return true;
						}
					}

					string uniqueMaTheLoai;
					bool isMaUnique = false;
					int attempts = 0;
					const int maxAttempts = 20;

					do
					{
						uniqueMaTheLoai = "TL" + GenerateRandomAlphaNumericCode(6);
						string checkMaQuery = "SELECT COUNT(*) FROM TheLoai WHERE MaTheLoai = @MaTheLoai COLLATE NOCASE";
						using (SQLiteCommand checkMaCmd = new SQLiteCommand(checkMaQuery, conn))
						{
							checkMaCmd.Parameters.AddWithValue("@MaTheLoai", uniqueMaTheLoai);
							long maCount = (long)checkMaCmd.ExecuteScalar();
							if (maCount == 0)
							{
								isMaUnique = true;
							}
						}
						attempts++;
						if (attempts > maxAttempts)
						{
							MessageBox.Show($"Không thể tạo Mã Thể Loại duy nhất cho '{tenTheLoaiInput}' sau {maxAttempts} lần thử.", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
							return false;
						}
					} while (!isMaUnique);

					string insertQuery = @"INSERT INTO TheLoai (MaTheLoai, TenTheLoai, MoTa, TrangThai)
                                         VALUES (@MaTheLoai, @TenTheLoai, @MoTa, @TrangThai)";
					using (SQLiteCommand insertCmd = new SQLiteCommand(insertQuery, conn))
					{
						insertCmd.Parameters.AddWithValue("@MaTheLoai", uniqueMaTheLoai);
						insertCmd.Parameters.AddWithValue("@TenTheLoai", tenTheLoaiInput);
						insertCmd.Parameters.AddWithValue("@MoTa", (object)DBNull.Value);
						insertCmd.Parameters.AddWithValue("@TrangThai", "Hoạt động");

						int rowsAffected = insertCmd.ExecuteNonQuery();
						if (rowsAffected > 0)
						{
							Console.WriteLine($"Đã tự động thêm thể loại mới: '{tenTheLoaiInput}' (Mã: {uniqueMaTheLoai}) vào bảng TheLoai.");
							return true;
						}
						else
						{
							Console.WriteLine($"Không thể thêm thể loại '{tenTheLoaiInput}' vào bảng TheLoai dù mã đã được tạo.");
							return false;
						}
					}
				}
				catch (SQLiteException ex)
				{
					MessageBox.Show($"Lỗi CSDL khi tự động thêm thể loại '{tenTheLoaiInput}': {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"SQLite Error adding category '{tenTheLoaiInput}': {ex.Message}");
					return false;
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi không mong muốn khi tự động thêm thể loại '{tenTheLoaiInput}': {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error adding category '{tenTheLoaiInput}': {ex.Message}");
					return false;
				}
			}
		}

		// --- Methods for NhaXuatBan ---
		public DataTable GetAllNhaXuatBan()
		{
			DataTable dataTable = new DataTable();
			string query = "SELECT ID, MaNhaXuatBan, TenNhaXuatBan, DiaChi, SoHieuXuatBan, Email, Website, NgayThanhLap, TrangThai FROM NhaXuatBan ORDER BY TenNhaXuatBan ASC";
			using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
			{
				try
				{
					sqlConnection.Open();
					SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, sqlConnection);
					adapter.Fill(dataTable);
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi lấy danh sách nhà xuất bản: {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error getting all publishers: {ex.Message}");
				}
			}
			return dataTable;
		}
		public List<string> GetAllTenNhaXuatBan()
		{
			List<string> tenNXBList = new List<string>();
			string query = "SELECT DISTINCT TenNhaXuatBan FROM NhaXuatBan WHERE TenNhaXuatBan IS NOT NULL AND TenNhaXuatBan != '' ORDER BY TenNhaXuatBan ASC";
			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						using (SQLiteDataReader reader = cmd.ExecuteReader())
						{
							while (reader.Read())
							{
								tenNXBList.Add(reader.GetString(0));
							}
						}
					}
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi lấy danh sách tên nhà xuất bản: {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error getting all publisher names: {ex.Message}");
				}
			}
			return tenNXBList;
		}

		public NhaXuatBan GetNhaXuatBanByName(string tenNXB)
		{
			NhaXuatBan nxb = null;
			string query = "SELECT * FROM NhaXuatBan WHERE TenNhaXuatBan = @TenNhaXuatBan COLLATE NOCASE LIMIT 1";
			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						cmd.Parameters.AddWithValue("@TenNhaXuatBan", tenNXB);
						using (SQLiteDataReader reader = cmd.ExecuteReader())
						{
							if (reader.Read())
							{
								nxb = new NhaXuatBan
								{
									ID = reader.GetInt32(reader.GetOrdinal("ID")),
									MaNhaXuatBan = reader.GetString(reader.GetOrdinal("MaNhaXuatBan")),
									TenNhaXuatBan = reader.GetString(reader.GetOrdinal("TenNhaXuatBan")),
									DiaChi = reader.IsDBNull(reader.GetOrdinal("DiaChi")) ? null : reader.GetString(reader.GetOrdinal("DiaChi")),
									SoHieuXuatBan = reader.IsDBNull(reader.GetOrdinal("SoHieuXuatBan")) ? null : reader.GetString(reader.GetOrdinal("SoHieuXuatBan")),
									Email = reader.IsDBNull(reader.GetOrdinal("Email")) ? null : reader.GetString(reader.GetOrdinal("Email")),
									Website = reader.IsDBNull(reader.GetOrdinal("Website")) ? null : reader.GetString(reader.GetOrdinal("Website")),
									NgayThanhLap = reader.IsDBNull(reader.GetOrdinal("NgayThanhLap")) ? null : reader.GetString(reader.GetOrdinal("NgayThanhLap")),
									TrangThai = reader.IsDBNull(reader.GetOrdinal("TrangThai")) ? null : reader.GetString(reader.GetOrdinal("TrangThai"))
								};
							}
						}
					}
				}
				catch (Exception ex)
				{
					Console.WriteLine($"Lỗi khi lấy nhà xuất bản theo tên '{tenNXB}': {ex.Message}");
				}
			}
			return nxb;
		}

		public bool InsertNhaXuatBan(NhaXuatBan nxb, bool generateMaNXB = true)
		{
			using (SQLiteConnection sqlConnection = Connection.GetSQLiteConnection())
			{
				try
				{
					sqlConnection.Open();

					if (!IsTenNhaXuatBanUnique(nxb.TenNhaXuatBan, sqlConnection))
					{
						Console.WriteLine($"Tên nhà xuất bản '{nxb.TenNhaXuatBan}' đã tồn tại.");
						return false;
					}

					if (generateMaNXB || string.IsNullOrWhiteSpace(nxb.MaNhaXuatBan))
					{
						string uniqueMaNXB;
						bool isUnique = false;
						int attempts = 0;
						const int maxAttempts = 20;
						do
						{
							uniqueMaNXB = GenerateRandomAlphaNumericCode(6, "NXB");
							isUnique = IsMaNhaXuatBanUnique(uniqueMaNXB, sqlConnection);
							attempts++;
							if (attempts > maxAttempts)
							{
								MessageBox.Show($"Không thể tạo Mã Nhà Xuất Bản duy nhất sau {maxAttempts} lần thử.", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
								return false;
							}
						} while (!isUnique);
						nxb.MaNhaXuatBan = uniqueMaNXB;
					}
					else
					{
						if (!IsMaNhaXuatBanUnique(nxb.MaNhaXuatBan, sqlConnection))
						{
							MessageBox.Show($"Mã Nhà Xuất Bản '{nxb.MaNhaXuatBan}' đã tồn tại.", "Lỗi Trùng Lặp", MessageBoxButtons.OK, MessageBoxIcon.Warning);
							return false;
						}
					}


					string query = @"INSERT INTO NhaXuatBan 
                                    (MaNhaXuatBan, TenNhaXuatBan, DiaChi, SoHieuXuatBan, Email, Website, NgayThanhLap, TrangThai)
                                    VALUES 
                                    (@MaNhaXuatBan, @TenNhaXuatBan, @DiaChi, @SoHieuXuatBan, @Email, @Website, @NgayThanhLap, @TrangThai)";
					using (SQLiteCommand cmd = new SQLiteCommand(query, sqlConnection))
					{
						cmd.Parameters.AddWithValue("@MaNhaXuatBan", nxb.MaNhaXuatBan);
						cmd.Parameters.AddWithValue("@TenNhaXuatBan", nxb.TenNhaXuatBan);
						cmd.Parameters.AddWithValue("@DiaChi", (object)nxb.DiaChi ?? DBNull.Value);
						cmd.Parameters.AddWithValue("@SoHieuXuatBan", (object)nxb.SoHieuXuatBan ?? DBNull.Value);
						cmd.Parameters.AddWithValue("@Email", (object)nxb.Email ?? DBNull.Value);
						cmd.Parameters.AddWithValue("@Website", (object)nxb.Website ?? DBNull.Value);
						cmd.Parameters.AddWithValue("@NgayThanhLap", string.IsNullOrEmpty(nxb.NgayThanhLap) ? (object)DBNull.Value : nxb.NgayThanhLap);
						cmd.Parameters.AddWithValue("@TrangThai", (object)nxb.TrangThai ?? "Hoạt động");

						int rowsAffected = cmd.ExecuteNonQuery();
						return rowsAffected > 0;
					}
				}
				catch (SQLiteException ex)
				{
					if (ex.ErrorCode == (int)SQLiteErrorCode.Constraint_Unique || ex.Message.ToUpper().Contains("UNIQUE CONSTRAINT FAILED"))
					{
						MessageBox.Show($"Lỗi: Mã nhà xuất bản '{nxb.MaNhaXuatBan}' hoặc Tên nhà xuất bản '{nxb.TenNhaXuatBan}' đã tồn tại.", "Lỗi Trùng Lặp", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					else
					{
						MessageBox.Show($"Lỗi CSDL khi thêm nhà xuất bản: {ex.Message} (Mã lỗi: {ex.ErrorCode})", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					}
					Console.WriteLine($"SQLite Error inserting publisher (Code {ex.ErrorCode}): {ex.Message}");
					return false;
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi không mong muốn khi thêm nhà xuất bản: {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error inserting publisher: {ex.Message}");
					return false;
				}
			}
		}

		public bool UpdateNhaXuatBan(NhaXuatBan nxb)
		{
			if (nxb == null || nxb.ID <= 0)
			{
				MessageBox.Show("Thông tin nhà xuất bản không hợp lệ để cập nhật (Thiếu ID).", "Lỗi Dữ liệu", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				return false;
			}

			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					if (!IsTenNhaXuatBanUnique(nxb.TenNhaXuatBan, conn, nxb.ID))
					{
						MessageBox.Show($"Tên nhà xuất bản '{nxb.TenNhaXuatBan}' đã được sử dụng bởi một nhà xuất bản khác.", "Lỗi Trùng Lặp", MessageBoxButtons.OK, MessageBoxIcon.Warning);
						return false;
					}

					string query = @"UPDATE NhaXuatBan
                                 SET TenNhaXuatBan = @TenNhaXuatBan,
                                     DiaChi = @DiaChi,
                                     SoHieuXuatBan = @SoHieuXuatBan,
                                     Email = @Email,
                                     Website = @Website,
                                     NgayThanhLap = @NgayThanhLap,
                                     TrangThai = @TrangThai
                                 WHERE ID = @ID";

					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						cmd.Parameters.AddWithValue("@ID", nxb.ID);
						cmd.Parameters.AddWithValue("@TenNhaXuatBan", nxb.TenNhaXuatBan ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@DiaChi", nxb.DiaChi ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@SoHieuXuatBan", nxb.SoHieuXuatBan ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@Email", nxb.Email ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@Website", nxb.Website ?? (object)DBNull.Value);
						cmd.Parameters.AddWithValue("@NgayThanhLap", string.IsNullOrEmpty(nxb.NgayThanhLap) ? (object)DBNull.Value : nxb.NgayThanhLap);
						cmd.Parameters.AddWithValue("@TrangThai", nxb.TrangThai ?? (object)DBNull.Value);

						int rowsAffected = cmd.ExecuteNonQuery();
						return rowsAffected > 0;
					}
				}
				catch (SQLiteException ex)
				{
					MessageBox.Show($"Lỗi CSDL khi cập nhật nhà xuất bản (ID: {nxb.ID}): {ex.Message} (Mã lỗi: {ex.ErrorCode})", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"DB Error Updating Publisher (ID: {nxb.ID}, Code: {ex.ErrorCode}): {ex.Message}");
					return false;
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi không mong muốn khi cập nhật nhà xuất bản (ID: {nxb.ID}): {ex.Message}", "Lỗi Hệ Thống", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error Updating Publisher (ID: {nxb.ID}): {ex.Message}");
					return false;
				}
			}
		}

		public bool AddNhaXuatBanIfNotExists(string tenNhaXuatBan)
		{
			if (string.IsNullOrWhiteSpace(tenNhaXuatBan))
			{
				return false;
			}

			NhaXuatBan existingNXB = GetNhaXuatBanByName(tenNhaXuatBan);
			if (existingNXB == null)
			{
				NhaXuatBan newNXB = new NhaXuatBan
				{
					TenNhaXuatBan = tenNhaXuatBan,
					TrangThai = "Hoạt động"
				};
				bool result = InsertNhaXuatBan(newNXB, true);
				if (result)
				{
					Console.WriteLine($"Đã tự động thêm nhà xuất bản mới: '{tenNhaXuatBan}' (Mã: {newNXB.MaNhaXuatBan})");
				}
				return result;
			}
			return true;
		}

		public DataTable GetBooksByPublisherName(string tenNhaXuatBan)
		{
			DataTable dataTable = new DataTable();
			string query = @"SELECT ID, Masach, TenSach, TheLoai, TacGia, SoLuong, NhaXuatBan, NamXuatBan, anh, pdf, tomtatnoidung
                           FROM quanlysach
                           WHERE NhaXuatBan = @TenNhaXuatBan COLLATE NOCASE
                           ORDER BY TenSach ASC";

			using (SQLiteConnection conn = Connection.GetSQLiteConnection())
			{
				try
				{
					conn.Open();
					SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn);
					adapter.SelectCommand.Parameters.AddWithValue("@TenNhaXuatBan", string.IsNullOrEmpty(tenNhaXuatBan) ? (object)DBNull.Value : tenNhaXuatBan);
					adapter.Fill(dataTable);
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Lỗi khi lấy sách theo nhà xuất bản '{tenNhaXuatBan}': {ex.Message}", "Lỗi CSDL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Console.WriteLine($"Error getting books by publisher name '{tenNhaXuatBan}': {ex.Message}");
				}
			}
			return dataTable;
		}

	}
}
