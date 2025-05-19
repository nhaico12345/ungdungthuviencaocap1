using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ungdungthuviencaocap
{
	class muontra
	{
		private int _ID;
		private string _tensach;
		private string _masinhvien;
		private string _hovaten;
		private string _loaiphieu;
		private DateTime _ngaymuon;
		private DateTime _ngaytra;
		private string _soluong;
		private string _masach; 

		public muontra()
		{
		}

		public muontra(int iD, string tensach, string masinhvien, string hovaten, string loaiphieu,
					   DateTime ngaymuon, DateTime ngaytra, string soluong, string masach)
		{
			ID = iD;
			Tensach = tensach;
			Masinhvien = masinhvien;
			Hovaten = hovaten;
			Loaiphieu = loaiphieu;
			Ngaymuon = ngaymuon;
			Ngaytra = ngaytra;
			Soluong = soluong;
			Masach = masach; 
		}

		public int ID { get => _ID; set => _ID = value; }
		public string Tensach { get => _tensach; set => _tensach = value; }
		public string Masinhvien { get => _masinhvien; set => _masinhvien = value; }
		public string Hovaten { get => _hovaten; set => _hovaten = value; }
		public string Loaiphieu { get => _loaiphieu; set => _loaiphieu = value; }
		public DateTime Ngaymuon { get => _ngaymuon; set => _ngaymuon = value; }
		public DateTime Ngaytra { get => _ngaytra; set => _ngaytra = value; }
		public string Soluong { get => _soluong; set => _soluong = value; }
		public string Masach { get => _masach; set => _masach = value; } 
	}
}
