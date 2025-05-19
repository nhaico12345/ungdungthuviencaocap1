// TacGia.cs
using System;

namespace ungdungthuviencaocap
{
	public class TacGia
	{
		// Thuộc tính tương ứng với các cột trong bảng tacgia
		public int ID { get; set; }
		public string MaTacGia { get; set; }
		public string HoTen { get; set; }
		public string HocHam { get; set; } // Có thể null
		public string HocVi { get; set; } // Có thể null
		public string DiaChi { get; set; } // Có thể null
		public string Email { get; set; } // Có thể null
		public string NgaySinh { get; set; } // Lưu dạng text, có thể null
		public string QuocTich { get; set; } // Có thể null
		public string Anh { get; set; } // Đường dẫn file ảnh, có thể null
		public string TieuSu { get; set; } // Có thể null
		public string TrangThai { get; set; } // Có thể null

		// Constructor mặc định (cần thiết cho một số thao tác)
		public TacGia() { }

		// Constructor để tạo đối tượng tác giả mới (thường dùng khi insert, không cần ID)
		public TacGia(string maTacGia, string hoTen, string hocHam = null, string hocVi = null,
					  string diaChi = null, string email = null, string ngaySinh = null,
					  string quocTich = null, string anh = null, string tieuSu = null, string trangThai = null)
		{
			MaTacGia = maTacGia;
			HoTen = hoTen;
			HocHam = hocHam;
			HocVi = hocVi;
			DiaChi = diaChi;
			Email = email;
			NgaySinh = ngaySinh;
			QuocTich = quocTich;
			Anh = anh;
			TieuSu = tieuSu;
			TrangThai = trangThai;
		}

		// Constructor đầy đủ (thường dùng khi đọc dữ liệu từ CSDL)
		public TacGia(int id, string maTacGia, string hoTen, string hocHam, string hocVi,
					  string diaChi, string email, string ngaySinh,
					  string quocTich, string anh, string tieuSu, string trangThai)
		{
			ID = id;
			MaTacGia = maTacGia;
			HoTen = hoTen;
			HocHam = hocHam;
			HocVi = hocVi;
			DiaChi = diaChi;
			Email = email;
			NgaySinh = ngaySinh;
			QuocTich = quocTich;
			Anh = anh;
			TieuSu = tieuSu;
			TrangThai = trangThai;
		}
	}
}