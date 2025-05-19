// TheLoai.cs
using System;

namespace ungdungthuviencaocap
{
	public class TheLoai
	{
		// Thuộc tính tương ứng với các cột trong bảng TheLoai
		public int ID { get; set; } // Khóa chính, tự động tăng
		public string MaTheLoai { get; set; } // Mã thể loại duy nhất
		public string TenTheLoai { get; set; }
		public string MoTa { get; set; } // Có thể null
		public string TrangThai { get; set; } // Có thể null

		// Constructor mặc định
		public TheLoai() { }

		// Constructor để tạo đối tượng thể loại mới (thường dùng khi insert, không cần ID)
		public TheLoai(string maTheLoai, string tenTheLoai, string moTa = null, string trangThai = null)
		{
			MaTheLoai = maTheLoai;
			TenTheLoai = tenTheLoai;
			MoTa = moTa;
			TrangThai = trangThai;
		}

		// Constructor đầy đủ (thường dùng khi đọc dữ liệu từ CSDL)
		public TheLoai(int id, string maTheLoai, string tenTheLoai, string moTa, string trangThai)
		{
			ID = id;
			MaTheLoai = maTheLoai;
			TenTheLoai = tenTheLoai;
			MoTa = moTa;
			TrangThai = trangThai;
		}
	}
}