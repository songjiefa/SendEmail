namespace SentEmail
{
	partial class Form1
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
			this.tb_to = new System.Windows.Forms.TextBox();
			this.bt_send = new System.Windows.Forms.Button();
			this.bt_connect = new System.Windows.Forms.Button();
			this.tb_emailAddress = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.tb_password = new System.Windows.Forms.TextBox();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.tb_url = new System.Windows.Forms.TextBox();
			this.label5 = new System.Windows.Forms.Label();
			this.tb_cc = new System.Windows.Forms.TextBox();
			this.lb_bcc = new System.Windows.Forms.Label();
			this.tb_bcc = new System.Windows.Forms.TextBox();
			this.tb_subject = new System.Windows.Forms.TextBox();
			this.label6 = new System.Windows.Forms.Label();
			this.tb_body = new System.Windows.Forms.TextBox();
			this.label7 = new System.Windows.Forms.Label();
			this.tb_loopTimes = new System.Windows.Forms.TextBox();
			this.label8 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// tb_to
			// 
			this.tb_to.Location = new System.Drawing.Point(138, 120);
			this.tb_to.Multiline = true;
			this.tb_to.Name = "tb_to";
			this.tb_to.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
			this.tb_to.Size = new System.Drawing.Size(350, 16);
			this.tb_to.TabIndex = 0;
			// 
			// bt_send
			// 
			this.bt_send.Location = new System.Drawing.Point(413, 413);
			this.bt_send.Name = "bt_send";
			this.bt_send.Size = new System.Drawing.Size(75, 23);
			this.bt_send.TabIndex = 1;
			this.bt_send.Text = "Send";
			this.bt_send.UseVisualStyleBackColor = true;
			this.bt_send.Click += new System.EventHandler(this.bt_send_Click);
			// 
			// bt_connect
			// 
			this.bt_connect.Location = new System.Drawing.Point(413, 65);
			this.bt_connect.Name = "bt_connect";
			this.bt_connect.Size = new System.Drawing.Size(75, 23);
			this.bt_connect.TabIndex = 2;
			this.bt_connect.Text = "Connect";
			this.bt_connect.UseVisualStyleBackColor = true;
			this.bt_connect.Click += new System.EventHandler(this.bt_connect_Click);
			// 
			// tb_emailAddress
			// 
			this.tb_emailAddress.Location = new System.Drawing.Point(138, 40);
			this.tb_emailAddress.Name = "tb_emailAddress";
			this.tb_emailAddress.Size = new System.Drawing.Size(214, 20);
			this.tb_emailAddress.TabIndex = 3;
			this.tb_emailAddress.Text = "will.fang@ucazhuhai.onmicrosoft.com";
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(56, 46);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(35, 13);
			this.label1.TabIndex = 4;
			this.label1.Text = "Email:";
			// 
			// tb_password
			// 
			this.tb_password.Location = new System.Drawing.Point(138, 65);
			this.tb_password.Name = "tb_password";
			this.tb_password.PasswordChar = '*';
			this.tb_password.Size = new System.Drawing.Size(214, 20);
			this.tb_password.TabIndex = 5;
			this.tb_password.Text = "I@mrootPWD";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(59, 65);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(56, 13);
			this.label2.TabIndex = 6;
			this.label2.Text = "Password:";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(68, 123);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(23, 13);
			this.label3.TabIndex = 7;
			this.label3.Text = "To:";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(59, 13);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(32, 13);
			this.label4.TabIndex = 8;
			this.label4.Text = "URL:";
			// 
			// tb_url
			// 
			this.tb_url.Location = new System.Drawing.Point(138, 6);
			this.tb_url.Name = "tb_url";
			this.tb_url.Size = new System.Drawing.Size(340, 20);
			this.tb_url.TabIndex = 9;
			this.tb_url.Text = "https://outlook.office365.com/EWS/Exchange.asmx";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(68, 150);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(23, 13);
			this.label5.TabIndex = 10;
			this.label5.Text = "Cc:";
			// 
			// tb_cc
			// 
			this.tb_cc.Location = new System.Drawing.Point(138, 150);
			this.tb_cc.Multiline = true;
			this.tb_cc.Name = "tb_cc";
			this.tb_cc.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
			this.tb_cc.Size = new System.Drawing.Size(350, 16);
			this.tb_cc.TabIndex = 11;
			// 
			// lb_bcc
			// 
			this.lb_bcc.AutoSize = true;
			this.lb_bcc.Location = new System.Drawing.Point(68, 176);
			this.lb_bcc.Name = "lb_bcc";
			this.lb_bcc.Size = new System.Drawing.Size(29, 13);
			this.lb_bcc.TabIndex = 12;
			this.lb_bcc.Text = "Bcc:";
			// 
			// tb_bcc
			// 
			this.tb_bcc.Location = new System.Drawing.Point(138, 176);
			this.tb_bcc.Multiline = true;
			this.tb_bcc.Name = "tb_bcc";
			this.tb_bcc.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
			this.tb_bcc.Size = new System.Drawing.Size(350, 16);
			this.tb_bcc.TabIndex = 13;
			// 
			// tb_subject
			// 
			this.tb_subject.Location = new System.Drawing.Point(138, 222);
			this.tb_subject.Name = "tb_subject";
			this.tb_subject.Size = new System.Drawing.Size(350, 20);
			this.tb_subject.TabIndex = 14;
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(68, 229);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(46, 13);
			this.label6.TabIndex = 15;
			this.label6.Text = "Subject:";
			// 
			// tb_body
			// 
			this.tb_body.Location = new System.Drawing.Point(138, 260);
			this.tb_body.Multiline = true;
			this.tb_body.Name = "tb_body";
			this.tb_body.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.tb_body.Size = new System.Drawing.Size(350, 122);
			this.tb_body.TabIndex = 16;
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(68, 263);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(34, 13);
			this.label7.TabIndex = 17;
			this.label7.Text = "Body:";
			// 
			// tb_loopTimes
			// 
			this.tb_loopTimes.Location = new System.Drawing.Point(138, 415);
			this.tb_loopTimes.Name = "tb_loopTimes";
			this.tb_loopTimes.Size = new System.Drawing.Size(61, 20);
			this.tb_loopTimes.TabIndex = 18;
			this.tb_loopTimes.Text = "1";
			// 
			// label8
			// 
			this.label8.AutoSize = true;
			this.label8.Location = new System.Drawing.Point(55, 418);
			this.label8.Name = "label8";
			this.label8.Size = new System.Drawing.Size(62, 13);
			this.label8.TabIndex = 19;
			this.label8.Text = "LoopTimes:";
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(558, 471);
			this.Controls.Add(this.label8);
			this.Controls.Add(this.tb_loopTimes);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.tb_body);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.tb_subject);
			this.Controls.Add(this.tb_bcc);
			this.Controls.Add(this.lb_bcc);
			this.Controls.Add(this.tb_cc);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.tb_url);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.tb_password);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.tb_emailAddress);
			this.Controls.Add(this.bt_connect);
			this.Controls.Add(this.bt_send);
			this.Controls.Add(this.tb_to);
			this.Name = "Form1";
			this.Text = "Form1";
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.TextBox tb_to;
		private System.Windows.Forms.Button bt_send;
		private System.Windows.Forms.Button bt_connect;
		private System.Windows.Forms.TextBox tb_emailAddress;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TextBox tb_password;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.TextBox tb_url;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.TextBox tb_cc;
		private System.Windows.Forms.Label lb_bcc;
		private System.Windows.Forms.TextBox tb_bcc;
		private System.Windows.Forms.TextBox tb_subject;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.TextBox tb_body;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.TextBox tb_loopTimes;
		private System.Windows.Forms.Label label8;
	}
}

