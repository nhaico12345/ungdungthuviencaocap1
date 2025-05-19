namespace ungdungthuviencaocap
{
	partial class huongdanquanly
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Windows.Forms.TreeNode treeNode1 = new System.Windows.Forms.TreeNode("Vấn đề về tài khoản và mật khẩu");
			System.Windows.Forms.TreeNode treeNode2 = new System.Windows.Forms.TreeNode("Sử dụng chức năng quản lý sách?");
			System.Windows.Forms.TreeNode treeNode3 = new System.Windows.Forms.TreeNode("Sử dụng chức năng mượn / trả sách?");
			System.Windows.Forms.TreeNode treeNode4 = new System.Windows.Forms.TreeNode("Giải đáp thắc mắc cho quản lý", new System.Windows.Forms.TreeNode[] {
            treeNode1,
            treeNode2,
            treeNode3});
			System.Windows.Forms.TreeNode treeNode5 = new System.Windows.Forms.TreeNode("Cách dùng mượn sách và hẹn trả");
			System.Windows.Forms.TreeNode treeNode6 = new System.Windows.Forms.TreeNode("Tôi muốn cập nhật thông tin");
			System.Windows.Forms.TreeNode treeNode7 = new System.Windows.Forms.TreeNode("Tôi muốn góp ý?");
			System.Windows.Forms.TreeNode treeNode8 = new System.Windows.Forms.TreeNode("Giải đáp thắc mắc cho người dùng", new System.Windows.Forms.TreeNode[] {
            treeNode5,
            treeNode6,
            treeNode7});
			System.Windows.Forms.TreeNode treeNode9 = new System.Windows.Forms.TreeNode("Tác giả ");
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(huongdanquanly));
			this.panel1 = new System.Windows.Forms.Panel();
			this.treeView1 = new System.Windows.Forms.TreeView();
			this.Listanh = new System.Windows.Forms.ImageList(this.components);
			this.panel2 = new System.Windows.Forms.Panel();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.panel1.SuspendLayout();
			this.panel2.SuspendLayout();
			this.SuspendLayout();
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.treeView1);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Left;
			this.panel1.Location = new System.Drawing.Point(0, 0);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(338, 761);
			this.panel1.TabIndex = 0;
			// 
			// treeView1
			// 
			this.treeView1.Cursor = System.Windows.Forms.Cursors.Hand;
			this.treeView1.Dock = System.Windows.Forms.DockStyle.Top;
			this.treeView1.Font = new System.Drawing.Font("Tahoma", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.treeView1.FullRowSelect = true;
			this.treeView1.ImageIndex = 0;
			this.treeView1.ImageList = this.Listanh;
			this.treeView1.Location = new System.Drawing.Point(0, 0);
			this.treeView1.Name = "treeView1";
			treeNode1.ImageIndex = 9;
			treeNode1.Name = "Node2";
			treeNode1.SelectedImageKey = "profile-user.png";
			treeNode1.Text = "Vấn đề về tài khoản và mật khẩu";
			treeNode2.ImageIndex = 5;
			treeNode2.Name = "Node4";
			treeNode2.SelectedImageKey = "book.png";
			treeNode2.Text = "Sử dụng chức năng quản lý sách?";
			treeNode3.ImageIndex = 6;
			treeNode3.Name = "Node5";
			treeNode3.SelectedImageKey = "borrow.png";
			treeNode3.Text = "Sử dụng chức năng mượn / trả sách?";
			treeNode4.Name = "Node_quanly";
			treeNode4.Text = "Giải đáp thắc mắc cho quản lý";
			treeNode5.ImageIndex = 4;
			treeNode5.Name = "Node6";
			treeNode5.SelectedImageKey = "book (1).png";
			treeNode5.Text = "Cách dùng mượn sách và hẹn trả";
			treeNode6.ImageIndex = 8;
			treeNode6.Name = "Node7";
			treeNode6.SelectedImageKey = "news.png";
			treeNode6.Text = "Tôi muốn cập nhật thông tin";
			treeNode7.ImageIndex = 7;
			treeNode7.Name = "Node9";
			treeNode7.SelectedImageKey = "comments.png";
			treeNode7.Text = "Tôi muốn góp ý?";
			treeNode8.ImageIndex = 1;
			treeNode8.Name = "Node_nguoidung";
			treeNode8.SelectedImageKey = "chat.png";
			treeNode8.Text = "Giải đáp thắc mắc cho người dùng";
			treeNode9.ImageKey = "writer.png";
			treeNode9.Name = "Node8";
			treeNode9.SelectedImageKey = "writer.png";
			treeNode9.Text = "Tác giả ";
			this.treeView1.Nodes.AddRange(new System.Windows.Forms.TreeNode[] {
            treeNode4,
            treeNode8,
            treeNode9});
			this.treeView1.SelectedImageIndex = 0;
			this.treeView1.Size = new System.Drawing.Size(338, 761);
			this.treeView1.TabIndex = 0;
			this.treeView1.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterSelect);
			// 
			// Listanh
			// 
			this.Listanh.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("Listanh.ImageStream")));
			this.Listanh.TransparentColor = System.Drawing.Color.Transparent;
			this.Listanh.Images.SetKeyName(0, "customer-support.png");
			this.Listanh.Images.SetKeyName(1, "chat.png");
			this.Listanh.Images.SetKeyName(2, "support.png");
			this.Listanh.Images.SetKeyName(3, "writer.png");
			this.Listanh.Images.SetKeyName(4, "book (1).png");
			this.Listanh.Images.SetKeyName(5, "book.png");
			this.Listanh.Images.SetKeyName(6, "borrow.png");
			this.Listanh.Images.SetKeyName(7, "comments.png");
			this.Listanh.Images.SetKeyName(8, "news.png");
			this.Listanh.Images.SetKeyName(9, "profile-user.png");
			this.Listanh.Images.SetKeyName(10, "studying.png");
			// 
			// panel2
			// 
			this.panel2.Controls.Add(this.label2);
			this.panel2.Controls.Add(this.label1);
			this.panel2.Dock = System.Windows.Forms.DockStyle.Right;
			this.panel2.Location = new System.Drawing.Point(547, 0);
			this.panel2.Name = "panel2";
			this.panel2.Size = new System.Drawing.Size(959, 761);
			this.panel2.TabIndex = 1;
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Font = new System.Drawing.Font("Tahoma", 10.2F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(27, 121);
			this.label1.MaximumSize = new System.Drawing.Size(900, 0);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(62, 21);
			this.label1.TabIndex = 0;
			this.label1.Text = "label1";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Font = new System.Drawing.Font("Tahoma", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label2.Location = new System.Drawing.Point(325, 42);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(284, 36);
			this.label2.TabIndex = 1;
			this.label2.Text = "Giải đáp thắc mắc";
			// 
			// huongdanquanly
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 16F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(1506, 761);
			this.Controls.Add(this.panel2);
			this.Controls.Add(this.panel1);
			this.Name = "huongdanquanly";
			this.Text = "Giải đáp thắc mắc";
			this.panel1.ResumeLayout(false);
			this.panel2.ResumeLayout(false);
			this.panel2.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.TreeView treeView1;
		private System.Windows.Forms.ImageList Listanh;
		private System.Windows.Forms.Panel panel2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
	}
}