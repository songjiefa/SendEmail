using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SentEmail
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}
		ExchangeService service = new ExchangeService();
		
		String EmailRegex = @"^[A-Za-z\d]+([-_.][A-Za-z\d]+)*@([A-Za-z\d]+[-.])+[A-Za-z\d]{2,4}$ ";
		
		private void bt_send_Click(object sender, EventArgs e)
		{
			var tos = tb_to.Text.Split(';');
			var ccs = tb_to.Text.Split(';');
			var bccs = tb_to.Text.Split(';');

			if (tb_to.Text != String.Empty && !Regex.IsMatch(tb_to.Text, EmailRegex))
			{
				MessageBox.Show(string.Format(
					"{0} has wrong email address format.",tb_to.Text));
			}
			if (tb_to.Text != String.Empty && !Regex.IsMatch(tb_cc.Text, EmailRegex))
			{
				MessageBox.Show(string.Format(
					"{0} has wrong email address format.", tb_cc.Text));
			}
			if (tb_to.Text != String.Empty && !Regex.IsMatch(tb_bcc.Text, EmailRegex))
			{
				MessageBox.Show(string.Format(
					"{0} has wrong email address format.", tb_bcc.Text));
			}


			EmailMessage message = new EmailMessage(service);
			message.Subject = "Interesting";

			message.Body = "The proposition has been considered.";

			message.ToRecipients.Add("SharedMailHead@ucazhuhai.onmicrosoft.com");
			message.ToRecipients.Add("SharedMailHead2@ucazhuhai.onmicrosoft.com");
			message.ToRecipients.Add("SharedMailHead3@ucazhuhai.onmicrosoft.com");
			message.Send();
		}

		private void bt_connect_Click(object sender, EventArgs e)
		{
			if (!Regex.IsMatch(tb_emailAddress.Text, EmailRegex))
			{
				MessageBox.Show(string.Format(
					"{0} has wrong email address format.", tb_emailAddress.Text));
			}
			// Connect by using the default credentials of the authenticated user.
			service.UseDefaultCredentials = true;

			// Connect by using the credentials of user1 at contoso.com.
			service.Credentials = new WebCredentials(tb_emailAddress.Text, tb_password.Text);

			// Connect by using the credentials of contoso/user1.
			//service.Credentials = new WebCredentials("user1", "password", "contoso");

			// Use Autodiscover to set the URL endpoint.
			//service.Credentials = new WebCredentials("user1", "password", "contoso");
			//service.AutodiscoverUrl("will.fang@ucazhuhai.onmicrosoft.com");
			service.Url = new Uri("");

			//service.begin
		}
	}
}
