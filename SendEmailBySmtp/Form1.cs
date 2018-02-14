using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;

namespace SendEmailBySmtp
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void bt_send_Click(object sender, EventArgs e)
		{
			SmtpMail oMail = new SmtpMail("TryIt");
			SmtpClient oSmtp = new SmtpClient();

			// Set sender email address, please change it to yours
			oMail.From = "test@emailarchitect.net";
			// Set recipient email address, please change it to yours
			oMail.To = "support@emailarchitect.net";

			// Set email subject
			oMail.Subject = "test email from c# project";

			// Set email body
			oMail.TextBody = "this is a test email sent from c# project, do not reply";

			// Your Exchange Server address
			SmtpServer oServer = new SmtpServer("exch.emailarchitect.net");

			// Set Exchange Web Service EWS - Exchange 2007/2010/2013/2016
			oServer.Protocol = ServerProtocol.ExchangeEWS;
			// User and password for Exchange authentication
			oServer.User = "test";
			oServer.Password = "testpassword";

			// By default, Exchange Web Service requires SSL connection
			oServer.ConnectType = SmtpConnectType.ConnectSSLAuto;

			try
			{
				Console.WriteLine("start to send email ...");
				oSmtp.Send(oServer, oMail);
				Console.WriteLine("email was sent successfully!");
			}
			catch (Exception ep)
			{
				Console.WriteLine("failed to send email with the following error:");
				Console.WriteLine(ep.Message);
			}
		}
	}
}
