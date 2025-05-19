using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ungdungthuviencaocap
{
    class henmuonsach
    {
		private int _ID;
		private string _masv;
		private string _hovaten;
		private string _tensach;	
		private DateTime _lichhen;
		private string _giohen;
		private string _loaihen;
		public henmuonsach()
		{
		}

		public henmuonsach(int ID,string masinhvien, string tensach, string hovaten, DateTime lichhen, string giohen, string loaihen)
		{
			_ID = ID;
			_masv = masinhvien;
			_hovaten = hovaten;
			_tensach = tensach;
			_lichhen = lichhen;
			_giohen = giohen;
			_loaihen = loaihen;
		}
		public int ID { get => _ID; set => _ID = value; }
		public string Masinhvien { get => _masv; set => _masv = value; }
		public string Tensach { get => _tensach; set => _tensach = value; }
		public string Hovaten { get => _hovaten; set => _hovaten = value; }
		public DateTime Lichhen { get => _lichhen; set => _lichhen = value; }
		public string Giohen { get => _giohen; set => _giohen = value; }
		public string Loaihen { get => _loaihen; set => _loaihen = value; }
	}
}
