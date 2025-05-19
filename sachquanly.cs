using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ungdungthuviencaocap
{
	class sachquanly
	{
		// Backing fields cho các thuộc tính hiện có
		private int _ID;
		private string _masach;
		private string _tensach;
		private string _theloai;
		private string _tacgia;
		private int _soluong;
		private string _nhaxuatban;
		private int _namxuatban;

		// *** THÊM Backing fields cho các thuộc tính mới ***
		private string _anh; // Đường dẫn file ảnh
		private string _pdf; // Đường dẫn file PDF
		private string _tomtatnoidung; // Nội dung tóm tắt

		// Constructor mặc định
		public sachquanly()
		{
			// Khởi tạo các giá trị mặc định nếu cần (ví dụ: chuỗi rỗng)
			Anh = null;
			Pdf = null;
			Tomtatnoidung = null;
		}

		// Constructor đầy đủ (bao gồm các trường mới)
		public sachquanly(int iD, string masach, string tensach, string theloai, string tacgia, int soluong, string nhaxuatban, int namxuatban, string anh, string pdf, string tomtatnoidung)
		{
			ID = iD;
			Masach = masach;
			Tensach = tensach;
			Theloai = theloai;
			Tacgia = tacgia;
			Soluong = soluong;
			Nhaxuatban = nhaxuatban;
			Namxuatban = namxuatban;
			// Gán giá trị cho các thuộc tính mới
			Anh = anh;
			Pdf = pdf;
			Tomtatnoidung = tomtatnoidung;
		}

		// Constructor cũ hơn (có thể giữ lại nếu vẫn dùng ở đâu đó, hoặc thêm giá trị null/mặc định cho trường mới)
		public sachquanly(int iD, string masach, string tensach, string theloai, string tacgia, int soluong, string nhaxuatban, int namxuatban)
			: this(iD, masach, tensach, theloai, tacgia, soluong, nhaxuatban, namxuatban, null, null, null) // Gọi constructor đầy đủ với giá trị null
		{
		}

		// Constructor rất cũ (không có masach - có thể giữ nếu logic tạo masach tự động vẫn cần)
		public sachquanly(int iD, string tensach, string theloai, string tacgia, int soluong, string nhaxuatban, int namxuatban)
			 : this(iD, null, tensach, theloai, tacgia, soluong, nhaxuatban, namxuatban, null, null, null) // Gọi constructor đầy đủ với masach và các trường mới là null
		{
		}


		// Thuộc tính (Properties) cho các trường hiện có
		public int ID { get => _ID; set => _ID = value; }
		public string Masach { get => _masach; set => _masach = value; }
		public string Tensach { get => _tensach; set => _tensach = value; }
		public string Theloai { get => _theloai; set => _theloai = value; }
		public string Tacgia { get => _tacgia; set => _tacgia = value; }
		public int Soluong { get => _soluong; set => _soluong = value; }
		public string Nhaxuatban { get => _nhaxuatban; set => _nhaxuatban = value; }
		public int Namxuatban { get => _namxuatban; set => _namxuatban = value; }

		// *** THÊM Thuộc tính cho các trường mới ***
		public string Anh { get => _anh; set => _anh = value; }
		public string Pdf { get => _pdf; set => _pdf = value; }
		public string Tomtatnoidung { get => _tomtatnoidung; set => _tomtatnoidung = value; }
	}
}