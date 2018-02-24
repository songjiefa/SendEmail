using Microsoft.Exchange.WebServices.Data;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SentEmail
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
			Control.CheckForIllegalCrossThreadCalls = false;
			LoadSettings();
		}

		const string settingFileSuffix = ".settings";

		private void LoadSettings()
		{
			lBSettings.Items.Clear();
			var settings = Directory.GetFiles(@".\", "*"+ settingFileSuffix, SearchOption.TopDirectoryOnly);
			if (settings.Count() > 0)
			{
				lBSettings.Items.AddRange(settings);
			}
		}

		ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010);

		//String EmailRegex = @"^[A-Za-z\d]+([-_.][A-Za-z\d]+)*@([A-Za-z\d]+[-.])+[A-Za-z\d]{2,4}$";
		String EmailRegex = @"[\w!#$%&'*+/=?^_`{|}~-]+(?:\.[\w!#$%&'*+/=?^_`{|}~-]+)*@(?:[\w](?:[\w-]*[\w])?\.)+[\w](?:[\w-]*[\w])?";
		private void bt_send_Click(object sender, EventArgs e)
		{
			//await SendEmail();
			var thread = new Thread(SendEmail);
			thread.Start();
		}

		private void SendEmail()
		{
			var tos = tb_to.Text.Split(';');
			var ccs = tb_cc.Text.Split(';');
			var bccs = tb_bcc.Text.Split(';');
			tos.ToList().ForEach(to =>
			{
				if (to != string.Empty && !Regex.IsMatch(to, EmailRegex))
				{
					MessageBox.Show(string.Format(
					"{0} has wrong email address format.", to));
				}
			});

			ccs.ToList().ForEach(cc =>
			{
				if (cc != string.Empty && !Regex.IsMatch(cc, EmailRegex))
				{
					MessageBox.Show(string.Format(
					"{0} has wrong email address format.", cc));
				}
			});

			bccs.ToList().ForEach(bcc =>
			{
				if (bcc != string.Empty && !Regex.IsMatch(bcc, EmailRegex))
				{
					MessageBox.Show(string.Format(
					"{0} has wrong email address format.", bcc));
				}
			});


			
			
			var random = new Random();
			var loopTime = int.Parse(tb_loopTimes.Text);
			progressBar1.Maximum = loopTime;
			for (int i = 0; i < loopTime; i++)
			{
				EmailMessage message = new EmailMessage(service);


				if (tos[0] != string.Empty) message.ToRecipients.AddRange(tos);
				if (ccs[0] != string.Empty) message.CcRecipients.AddRange(ccs);
				if (bccs[0] != string.Empty) message.ToRecipients.AddRange(bccs);
				message.Subject = tb_subject.Text + i.ToString("00000");

				message.Body = tb_body.Text + i.ToString("00000");
				//message.Attachments.AddFileAttachment("");
				var randomNum = random.Next(1, 100);
				if (randomNum < float.Parse(tb_attachmentRate.Text) * 100 &&
					tb_attachments.Tag != null)
				{
					message.Attachments.Clear();
					foreach (var fileNmae in (string[])tb_attachments.Tag)
					{
						message.Attachments.AddFileAttachment(fileNmae);
					}
				}
				
				message.Send();
				progressBar1.Value = i + 1;
				//Thread.Sleep(TimeSpan.FromSeconds(5));
			}
		}

		private void bt_connect_Click(object sender, EventArgs e)
		{
			//Trust all certificates
			System.Net.ServicePointManager.ServerCertificateValidationCallback =
		((sender1, certificate, chain, sslPolicyErrors) => true);

			// trust sender
			System.Net.ServicePointManager.ServerCertificateValidationCallback
							= ((sender1, cert, chain, errors) =>
							cert.Subject.Contains("CN"));

			// validate cert by calling a function
			ServicePointManager.ServerCertificateValidationCallback += new RemoteCertificateValidationCallback(ValidateRemoteCertificate);
			if (!Regex.IsMatch(tb_emailAddress.Text, EmailRegex))
			{
				MessageBox.Show(string.Format(
					"{0} has wrong email address format.", tb_emailAddress.Text));
			}
			// Connect by using the default credentials of the authenticated user.
			service.UseDefaultCredentials = true;

			// Connect by using the credentials of user1 at contoso.com.
			if (tb_url.Text.Contains("office365.com"))
			{
				service.Credentials = new WebCredentials(tb_emailAddress.Text, tb_password.Text);
			}
			else
			{
				service.Credentials = new WebCredentials(tb_emailAddress.Text, tb_password.Text);
				//service.Credentials = new NetworkCredential(tb_emailAddress.Text, tb_password.Text);
				//service.Credentials = DefaultCredential
			}

			// Connect by using the credentials of contoso/user1.
			//service.Credentials = new WebCredentials("user1", "password", "contoso");

			// Use Autodiscover to set the URL endpoint.
			//service.Credentials = new WebCredentials("user1", "password", "contoso");
			//service.AutodiscoverUrl("will.fang@ucazhuhai.onmicrosoft.com");
			service.Url = new Uri(tb_url.Text);

			//service.begin
		}



		// callback used to validate the certificate in an SSL conversation
		private static bool ValidateRemoteCertificate(object sender, X509Certificate cert, X509Chain chain, SslPolicyErrors policyErrors)
		{
			bool result = false;
			if (cert.Subject.ToUpper().Contains("CN"))
			{
				result = true;
			}

			return result;
		}

		private void tb_loopTimes_KeyPress(object sender, KeyPressEventArgs e)
		{
			//阻止从键盘输入键
			e.Handled = true;
			if ((e.KeyChar >= '0' && e.KeyChar <= '9') || e.KeyChar == '\b')
			{
				e.Handled = false;
			}
		}

		private void tb_attachmentRate_KeyPress(object sender, KeyPressEventArgs e)
		{
			e.Handled = true;
			if ((e.KeyChar >= '0' && e.KeyChar <= '9') || 
				(e.KeyChar == '.' && !tb_attachmentRate.Text.Contains('.')) ||
				e.KeyChar == '\b')
			{
				e.Handled = false;
			}
		}

		private void BtSelect_Click(object sender, EventArgs e)
		{
			if( openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				tb_attachments.Text = string.Join("\r\n", openFileDialog1.FileNames);
				tb_attachments.Tag = openFileDialog1.FileNames;
			}

			
		}

		private void btSave_Click(object sender, EventArgs e)
		{
			EwsInfo ewsinfo = new EwsInfo()
			{
				URL = tb_url.Text,
				LoginEmail = tb_emailAddress.Text,
				Password = tb_password.Text,
				To = tb_to.Text,
				Cc = tb_cc.Text,
				Bcc = tb_bcc.Text,
				Subject = tb_subject.Text,
				Body = tb_body.Text,
				Attachments = tb_attachments.Text,
				LoopTimes = int.Parse(tb_loopTimes.Text),
				AttachmentRate = float.Parse(tb_attachmentRate.Text)
			};

			if (tbSettingName.Text != String.Empty)
			{
				var json = JsonConvert.SerializeObject(ewsinfo);
				var file = string.Format(@"{0}\{1}{2}",
					Path.GetDirectoryName(Application.ExecutablePath), tbSettingName.Text, settingFileSuffix);
				File.WriteAllText(file, json);
				LoadSettings();
			}
		}

		private void btLoad_Click(object sender, EventArgs e)
		{
			if(lBSettings.SelectedItem != null)
			{
				var ewsInfo = JsonConvert.DeserializeObject<EwsInfo>(
					File.ReadAllText(lBSettings.SelectedItem.ToString()));

				loadSettings(ewsInfo);
			}
		}

		private void loadSettings(EwsInfo i_ewsInfo)
		{
			tb_url.Text = i_ewsInfo.URL;
			tb_emailAddress.Text = i_ewsInfo.LoginEmail;
			tb_password.Text = i_ewsInfo.Password;
			tb_to.Text = i_ewsInfo.To;
			tb_cc.Text = i_ewsInfo.Cc;
			tb_bcc.Text = i_ewsInfo.Bcc;
			tb_subject.Text = i_ewsInfo.Subject;
			tb_body.Text = i_ewsInfo.Body;
			tb_attachments.Text = i_ewsInfo.Attachments;
			tb_loopTimes.Text = i_ewsInfo.LoopTimes.ToString();
			tb_attachmentRate.Text = i_ewsInfo.AttachmentRate.ToString();
		}

		private void btDelete_Click(object sender, EventArgs e)
		{
			if(lBSettings.SelectedItem != null)
			{
				File.Delete(lBSettings.SelectedItem.ToString());
				LoadSettings();
			}
		}
	}
}
