using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Windows.Forms.DataVisualization.Charting;
using System.Data.SQLite;
using System.IO;
namespace ungdungthuviencaocap
{
	public partial class bieudothongke: DevExpress.XtraEditors.XtraForm
	{
		public bieudothongke()
		{
            InitializeComponent();
			SetupChart();
            // Nạp dữ liệu từ SQLite vào biểu đồ
            LoadChartData();
			caidatchart();
			LoadChartmuontra();
		}
		private void SetupChart()
		{
			// Xóa các series (nếu có) và tạo series mới cho biểu đồ tròn
			chartSach.Series.Clear();
			Series series = new Series("SachSeries");
			series.ChartType = SeriesChartType.Pie;
			series.IsValueShownAsLabel = true; // Hiển thị giá trị trên biểu đồ
			chartSach.Series.Add(series);

			// Tùy chỉnh thêm (title, legend, v.v) nếu cần
			chartSach.Titles.Clear();
			chartSach.Titles.Add("Biểu đồ thống kê số lượng sách theo tên sách");
		}

		private void LoadChartData()
		{
			try
			{
				using (SQLiteConnection conn = Connection.GetSQLiteConnection())
				{
					conn.Open();
					// Lấy tên sách và số lượng từ bảng quanlysach
					string query = "SELECT TenSach, SoLuong FROM quanlysach";
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						using (SQLiteDataReader reader = cmd.ExecuteReader())
						{
							while (reader.Read())
							{
								string tenSach = reader["TenSach"].ToString();
								int soLuong = int.Parse(reader["SoLuong"].ToString());

								// Thêm điểm dữ liệu vào series
								chartSach.Series["SachSeries"].Points.AddXY(tenSach, soLuong);
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Có lỗi khi kết nối hoặc lấy dữ liệu: " + ex.Message);
			}
		}
		private void caidatchart()
		{
			// Xóa series cũ và tạo mới
			chartMuontra.Series.Clear();
			Series series = new Series("MuonTraSeries");
			series.ChartType = SeriesChartType.Pie;  // Chọn kiểu biểu đồ tròn
			series.IsValueShownAsLabel = true;         // Hiển thị giá trị của SoLuong trên biểu đồ
			chartMuontra.Series.Add(series);

			// Thêm tiêu đề cho biểu đồ
			chartMuontra.Titles.Clear();
			chartMuontra.Titles.Add("Biểu đồ thống kê số lượng mượn sách");
		}

		private void LoadChartmuontra()
		{
			try
			{
				using (SQLiteConnection conn = Connection.GetSQLiteConnection())
				{
					conn.Open();
					// Lấy dữ liệu TenSach, HoVaTen và SoLuong từ bảng muontrasach
					string query = "SELECT TenSach, HoVaTen, SoLuong FROM muontrasach";
					using (SQLiteCommand cmd = new SQLiteCommand(query, conn))
					{
						using (SQLiteDataReader reader = cmd.ExecuteReader())
						{
							while (reader.Read())
							{
								string tenSach = reader["TenSach"].ToString();
								string hoVaTen = reader["HoVaTen"].ToString();
								int soLuong = int.Parse(reader["SoLuong"].ToString());

								// Tạo nhãn kết hợp giữa TenSach và HoVaTen (có thể thay đổi định dạng theo ý muốn)
								string label = $"{tenSach} - {hoVaTen}";

								// Thêm dữ liệu vào series của biểu đồ: nhãn và giá trị SoLuong
								chartMuontra.Series["MuonTraSeries"].Points.AddXY(label, soLuong);
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Có lỗi khi kết nối hoặc lấy dữ liệu: " + ex.Message);
			}
		}
	}
}