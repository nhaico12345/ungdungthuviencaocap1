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
using Microsoft.Reporting.WinForms;

namespace ungdungthuviencaocap
{
	public partial class formReportmuontrasach: DevExpress.XtraEditors.XtraForm
	{
        public formReportmuontrasach()
		{
            InitializeComponent();
		}

		Modify Modify = new Modify();
		private void formReportmuontrasach_Load(object sender, EventArgs e)
		{
			try
			{
				reportViewer1.LocalReport.ReportEmbeddedResource = "ungdungthuviencaocap.Report2.rdlc";
				ReportDataSource reportDataSource = new ReportDataSource();
				reportDataSource.Name = "muontrasach";
				reportDataSource.Value = Modify.getAllphieu();
				reportViewer1.LocalReport.DataSources.Add(reportDataSource);
				this.reportViewer1.RefreshReport();
			}
			catch (Exception ex)
			{
				MessageBox.Show("Lỗi khi nạp báo cáo: " + ex.Message);
			}
		}
	}
}