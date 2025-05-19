using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ungdungthuviencaocap
{
	class TaiKhoan
	{
		private string tenTaiKhoan;
		private string matKhau;

		public TaiKhoan()
		{
		}

		public TaiKhoan(string maSinhVien, string tenTaiKhoan, string matKhau, string email,
					string hoVaTen, string avatar, string quyen, string diaChi,
					string soDienThoai, string ngaySinh, string gioiTinh)
		{
			this.MaSinhVien = maSinhVien;
			this.TenTaiKhoan = tenTaiKhoan;
			this.MatKhau = matKhau;
			this.Email = email;
			this.HoVaTen = hoVaTen;
			this.Avatar = avatar;
			this.Quyen = quyen;
			this.DiaChi = diaChi;
			this.SoDienThoai = soDienThoai;
			this.NgaySinh = ngaySinh;
			this.GioiTinh = gioiTinh;
		}

		public string TenTaiKhoan { get => tenTaiKhoan; set => tenTaiKhoan = value; }
		public string MatKhau { get => matKhau; set => matKhau = value; }
		public string MaSinhVien { get; set; }
		public string Email { get; set; }
		public string HoVaTen { get; set; }
		public string Avatar { get; set; }
		public string Quyen { get; set; }
		public string DiaChi { get; set; }
		public string SoDienThoai { get; set; }
		public string NgaySinh { get; set; }
		public string GioiTinh { get; set; }
		public string TrangThai { get; internal set; }
	}
}
