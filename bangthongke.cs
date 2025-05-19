using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraBars;
using DevExpress.XtraBars.FluentDesignSystem;

namespace ungdungthuviencaocap
{
	public partial class bangthongke : DevExpress.XtraBars.Ribbon.RibbonForm
	{
		public bangthongke()
		{
			try
			{
				InitializeComponent();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khởi tạo form: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void barButtonItem1_ItemClick(object sender, ItemClickEventArgs e)
		{
			try
			{
				Modify modify = new Modify();
				SaveFileDialog saveFileDialog = new SaveFileDialog();
				saveFileDialog.Filter = "Excel Files|*.xlsx";
				saveFileDialog.Title = "Save an Excel File";
				if (saveFileDialog.ShowDialog() == DialogResult.OK)
				{
					string filePath = saveFileDialog.FileName;
					modify.ExportToExcelquanlysach(filePath);
					MessageBox.Show("Xuất file Excel thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi xuất file Excel quản lý sách: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void barButtonItem2_ItemClick(object sender, ItemClickEventArgs e)
		{
			try
			{
				Modify modify = new Modify();
				SaveFileDialog saveFileDialog = new SaveFileDialog();
				saveFileDialog.Filter = "Excel Files|*.xlsx";
				saveFileDialog.Title = "Save an Excel File";
				if (saveFileDialog.ShowDialog() == DialogResult.OK)
				{
					string filePath = saveFileDialog.FileName;
					modify.ExportToExcelmuontrasach(filePath);
					MessageBox.Show("Xuất file Excel thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi xuất file Excel mượn trả sách: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void barButtonItem3_ItemClick(object sender, ItemClickEventArgs e)
		{
			try
			{
				Modify modify = new Modify();
				SaveFileDialog saveFileDialog = new SaveFileDialog();
				saveFileDialog.Filter = "Excel Files|*.xlsx";
				saveFileDialog.Title = "Save an Excel File";
				if (saveFileDialog.ShowDialog() == DialogResult.OK)
				{
					string filePath = saveFileDialog.FileName;
					modify.ExportToExcel(filePath);
					MessageBox.Show("Xuất file Excel thành công!", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi xuất tất cả dữ liệu ra Excel: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void barButtonItem4_ItemClick(object sender, ItemClickEventArgs e)
		{
			try
			{
				openform(typeof(danhsachsach));
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form danh sách sách: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void barButtonItem5_ItemClick(object sender, ItemClickEventArgs e)
		{
			try
			{
				openform(typeof(danhsachmuontra));
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form danh sách mượn trả: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		void openform(Type typeform)
		{
			try
			{
				foreach (Form form in this.MdiChildren)
				{
					if (form.GetType() == typeform)
					{
						form.Activate();
						return;
					}
				}
				Form f = (Form)Activator.CreateInstance(typeform);
				f.MdiParent = this;
				f.Show();
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void barButtonItem6_ItemClick(object sender, ItemClickEventArgs e)
		{
			try
			{
				openform(typeof(bieudothongke));
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form biểu đồ thống kê: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void barButtonItem8_ItemClick(object sender, ItemClickEventArgs e)
		{
			try
			{
				MessageBox.Show("Không xuất file dưới dạng PDF để tránh lỗi nhé", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
				openform(typeof(formReportquanlysach));
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form báo cáo quản lý sách: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void barButtonItem9_ItemClick(object sender, ItemClickEventArgs e)
		{
			try
			{
				MessageBox.Show("Không xuất file dưới dạng PDF để tránh lỗi nhé", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
				openform(typeof(formReportmuontrasach));
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form báo cáo mượn trả sách: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void barButtonItem10_ItemClick(object sender, ItemClickEventArgs e)
		{
			try
			{
				openform(typeof(muontranhieunhat));
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form thống kê mượn trả nhiều nhất: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void barButtonItem11_ItemClick(object sender, ItemClickEventArgs e)
		{
			try
			{
				openform(typeof(thongkethongtinsach));
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Lỗi khi mở form thống kê thông tin sách: {ex.Message}", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}
