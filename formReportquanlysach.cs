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
	public partial class formReportquanlysach: DevExpress.XtraEditors.XtraForm
	{
        public formReportquanlysach()
		{
            InitializeComponent();
		}
		Modify Modify = new Modify();
		private void formReportquanlysach_Load(object sender, EventArgs e)
		{
			try
			{
			    reportViewer1.LocalReport.ReportEmbeddedResource = "ungdungthuviencaocap.Report1.rdlc";
				ReportDataSource reportDataSource = new ReportDataSource();
				reportDataSource.Name = "quanlysach";
				reportDataSource.Value = Modify.getAllbook();
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