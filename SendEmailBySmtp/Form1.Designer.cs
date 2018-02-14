namespace SendEmailBySmtp
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
			this.bt_send = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// bt_send
			// 
			this.bt_send.Location = new System.Drawing.Point(162, 78);
			this.bt_send.Name = "bt_send";
			this.bt_send.Size = new System.Drawing.Size(75, 23);
			this.bt_send.TabIndex = 0;
			this.bt_send.Text = "Send";
			this.bt_send.UseVisualStyleBackColor = true;
			this.bt_send.Click += new System.EventHandler(this.bt_send_Click);
			// 
			// Form1
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(284, 261);
			this.Controls.Add(this.bt_send);
			this.Name = "Form1";
			this.Text = "Form1";
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button bt_send;
	}
}

