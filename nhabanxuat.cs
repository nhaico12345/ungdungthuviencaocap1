// NhaXuatBan.cs
using System;

namespace ungdungthuviencaocap
{
	public class NhaXuatBan
	{
		public int ID { get; set; }
		public string MaNhaXuatBan { get; set; }
		public string TenNhaXuatBan { get; set; }
		public string DiaChi { get; set; }
		public string SoHieuXuatBan { get; set; }
		public string Email { get; set; }
		public string Website { get; set; }
		public string NgayThanhLap { get; set; } // Lưu dạng TEXT YYYY-MM-DD
		public string TrangThai { get; set; }

		// Constructor mặc định
		public NhaXuatBan() { }

		// Constructor đầy đủ
		public NhaXuatBan(int id, string maNhaXuatBan, string tenNhaXuatBan, string diaChi,
						  string soHieuXuatBan, string email, string website,
						  string ngayThanhLap, string trangThai)
		{
			ID = id;
			MaNhaXuatBan = maNhaXuatBan;
			TenNhaXuatBan = tenNhaXuatBan;
			DiaChi = diaChi;
			SoHieuXuatBan = soHieuXuatBan;
			Email = email;
			Website = website;
			NgayThanhLap = ngayThanhLap;
			TrangThai = trangThai;
		}
	}
}
