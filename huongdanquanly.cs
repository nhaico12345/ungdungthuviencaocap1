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

namespace ungdungthuviencaocap
{
	public partial class huongdanquanly : DevExpress.XtraEditors.XtraForm
	{
		public huongdanquanly()
		{
			InitializeComponent();
			label1.Text = "";
			treeView1.AfterSelect += treeView1_AfterSelect;
		}

		private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
		{
			switch (e.Node.Name)
			{
				case "Node2":
					label1.Text = "- Tài khoản của quản lý được cấp riêng, không được tạo theo cách bình thường để đảm bảo về vấn đề bảo mật của dữ liệu cũng như ứng dụng\n- Ứng dụng được phân quyền giữa quản lý và người dùng\n- Muốn truy cập và sử dụng quyền quản lý vui lòng liên hệ tới tác giả để được hỗ trợ.";
					break;
				case "Node4":
					label1.Text = "Chức năng quản lý sách bao gồm các chức năng sau:\n  - Thêm sách\n  - Xóa sách\n  - Sửa thông tin\n  - Nhập lại\n  - Tra sách\n  - Nhập dữ liệu bằng file Excel\nPhía trên là từng chức năng, bây giờ sẽ hướng dãn sử dụng từng chức năng:\n  - Sau khi đã nhập đầy đủ thông tin vào ô thì hãy ấn thêm sách\n  - Sử dụng nhập lại để xóa toàn bộ thông tin trong ô và nhập lại từ đầu\n  - Để sử dụng sửa thông tin thì bạn hãy nhập vào từng ô thông tin đã sửa và chọn dòng cần sửa thông tin phía bên dưới và ấn sửa, hệ thống sẽ tự cập nhật thông tin đã sửa, bạn hãy kiểm tra thông tin khi đã sửa xong xem đã chính xác dòng cần sửa chưa nhé\n  - Nhập tên sách cần tra vào ô và ấn tra sách sẽ hiện ra tên sách gần đúng hoặc đúng có trong dữ liệu\n  - Nếu bạn đã có sẵn file Excel thì chỉ cần ấn vào nhập file và tìm tới file cần nhật dữ liệu và ấn ok thì sẽ tự nhập dữ liệu vào trong hệ thống\n  - Nếu bạn muốn xóa sách có trong dữ liệu của hệ thống thì chỉ cần chọn dòng mang tên sách cần xóa thông tin và ấn xóa sách thì sách sẽ tự xóa trong dữ liệu hệ thống\n\nLưu ý: Nếu sau khi tìm kiếm tên sách và muốn hiện lại toàn bộ sách thì chỉ cần xóa tên sách đang tìm kiếm và ấn vào ấn tra sách để hiện lại toàn bộ dữ liêu trong bảng hoặc bạn có thể ấn vào mục quản lý sách bên phần chức năng phía bên trái để hiện lại bảng";
					break;
				case "Node5":
					label1.Text = "Chức năng mượn/ trả sách bao gồm các chức năng sau:\n   - Tạo phiếu mượn\n  - Sửa\n  - Kiểm tra\n  - Trả sách\nPhía trên là từng chức năng, bây giờ sẽ hướng dãn sử dụng từng chức năng:\n  - Sau khi đã nhập đủ thông tin vào ô thì ấn tạo phiếu mượn để tạo phiếu mượn sách cho người mượn, phiếu sẽ xuất hiện dưới bảng bên dưới\n  - Bạn nhập đầy đủ thông tin cần sửa vào các ô trên và chọn dòng cần sửa phía dưới, xong các bước trên ấn nút sửa thì dữ liệu sửa sẽ được cập nhật và dữ liệu ứng dụng\n  - Nhập mã sinh viên và ấn kiểm tra để hiện ra thông tin các sách đã mượn phía dưới sau khi đã xcá nhận là người dùng trả sách chỉ cần ấn vào nút trả sách để xóa đi dữ liệu người trả\n\nLưu ý: Nếu sau khi tìm kiếm mã sinh viên và muốn hiện lại toàn bộ phiếu mượn sách thì chỉ cần xóa mã sinh viên đang tìm kiếm và ấn vào ấn trả sách để hiện lại toàn bộ dữ liêu trong bảng hoặc bạn có thể ấn vào mục Mượn / trả sách bên phần chức năng phía bên trái để hiện lại bảng";
					break;
				case "Node6":
					label1.Text = "Chức năng đặt lịch hẹn mượn / trả sách thư viện là: Đặt lịch hẹn\n Sau khi điền đầy đủ thông tin lịch hẹn phía trên chỉ cần ấn vào Đặt lịch hẹn để tạo phiếu đặt lịch. Hãy nhớ đến đúng ngày / giờ đã hẹn nhé";
					break;
				case "Node7":
					label1.Text = "Chức năng cập nhật thông tin đang trong giai đoạn phát triển, bạn vui lòng đợi nhé";
					break;
				case "Node9":
					label1.Text = "Bạn muốn góp ý thì vui lòng ấn vào phần tác giả để có thông tin liên hệ nhé.";
					break;
				case "Node8":
					label1.Text = "Tác giả: Trần Đình Đăng\nQuyền sở hữu trí tuệ và sở hữu ứng dụng thuộc về: Trần Đình Đăng - K16 CNTT\nMọi thắc mắc và góp ý xin vui lòng liên hệ qua:\n  - Email: oosp0305@gmail.com\n  - Số điện thoại: 0962938324";
					break;
				default:
					label1.Text = "";
					break;
			}
		}
	}
}