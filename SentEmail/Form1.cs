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
		const int threadCount = 10;
		const string settingFileSuffix = ".settings";
		string password = "I@mroot";
		static ReaderWriterLockSlim LogWriteLock = new ReaderWriterLockSlim();
		private List<string> m_urls = new List<string>();

		private void LoadSettings()
		{
			lBSettings.Items.Clear();
			var settings = Directory.GetFiles(@".\", "*"+ settingFileSuffix, SearchOption.TopDirectoryOnly);
			if (settings.Count() > 0)
			{
				lBSettings.Items.AddRange(settings);
			}
		}

		//ExchangeService Service = new ExchangeService(ExchangeVersion.Exchange2010);



		//String EmailRegex = @"^[A-Za-z\d]+([-_.][A-Za-z\d]+)*@([A-Za-z\d]+[-.])+[A-Za-z\d]{2,4}$";
		String EmailRegex = @"[\w!#$%&'*+/=?^_`{|}~-]+(?:\.[\w!#$%&'*+/=?^_`{|}~-]+)*@(?:[\w](?:[\w-]*[\w])?\.)+[\w](?:[\w-]*[\w])?";

		public ExchangeService Service
		{
			get
			{
				return new ExchangeService(ExchangeVersion.Exchange2010);
			}

			
		}

		private void bt_send_Click(object sender, EventArgs e)
		{
			//await SendEmail();
			var thread = new Thread(SendEmail);
			thread.Start();
		}

		private void SendEmail()
		{
			tb_to.Text = SetRecipients(tb_to.Text);
			tb_cc.Text = SetRecipients(tb_cc.Text);
			tb_bcc.Text = SetRecipients(tb_bcc.Text);

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


			if (cbJustGenerateSendItems.Checked)
			{
				
				var result = GenerateSendItems(tos, ccs, bccs, random, loopTime);
				while (!result.IsCompleted)
				{
					Thread.Sleep(1000 * 20);
				};

				this.Text = "Done";
				return;
			}

			if (cbSendByTos.Checked)
			{
				progressBar1.Maximum = tos.Count() * tos.Count() * 2 + tos.Count() * 2;
				var result =  CreateFolder(tos);
				while (!result.IsCompleted)
				{
					Thread.Sleep(1000 * 20);
				};

				result = SendByTos(tos, ccs, bccs, random, loopTime, true);
				while (!result.IsCompleted)
				{
					Thread.Sleep(1000 * 20);
				};

				result = MoveInboxEmailToCustomFolder(tos, "Custom Folder", 500);
				while (!result.IsCompleted)
				{
					Thread.Sleep(1000 * 20);
				};

				result = SendByTos(tos, ccs, bccs, random, loopTime, false);
				while (!result.IsCompleted)
				{
					Thread.Sleep(1000 * 20);
				};

				MessageBox.Show("Done!");
				return;
			}


			if (cbIsSendIndividual.Checked)
			{
				SendIndividual(tos, ccs, bccs, random, loopTime);
				MessageBox.Show("Done!");
				return;
			}
			
			SendEmail(tos, ccs, bccs, random, loopTime);
			
		}

		private void SendEmail(string[] tos, string[] ccs, string[] bccs, Random random, int loopTime)
		{
			progressBar1.Maximum = loopTime;
			for (int i = 0; i < loopTime; i++)
			{
				Service.Url = new Uri(m_urls[i % m_urls.Count]);

				EmailMessage message = new EmailMessage(Service);


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

		private void SendIndividual(string[] tos, string[] ccs, string[] bccs, Random random, int loopTime)
		{
			progressBar1.Maximum = loopTime * tos.Count();
			foreach (var to in tos)
			{

				for (int i = 0; i < loopTime; i++)
				{
					Service.Url = new Uri(m_urls[i % m_urls.Count]);
					EmailMessage message = new EmailMessage(Service);

					message.ToRecipients.Add(to);
					if (ccs[0] != string.Empty) message.CcRecipients.AddRange(ccs);
					if (bccs[0] != string.Empty) message.ToRecipients.AddRange(bccs);
					message.Subject = tb_subject.Text + to + i.ToString("00000");

					message.Body = tb_body.Text + to + i.ToString("00000");
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
					progressBar1.Value++;
					//Thread.Sleep(TimeSpan.FromSeconds(5));
				}
			}
		}

		private string SetRecipients(string i_partipient)
		{
			var emails = new List<string>();
			Regex pattern = new Regex(@"(?<prefix>.+);(?<startIndex>\d+)-(?<endIndex>\d+)");
			Match match = pattern.Match(i_partipient);
			if (match.Success)
			{
				var prifix = match.Groups["prefix"].Value;
				var startIndex = int.Parse(match.Groups["startIndex"].Value);
				var endIndex = int.Parse(match.Groups["endIndex"].Value);
				for (int i = startIndex; i <= endIndex; i++)
				{
					var email = 
						string.Format(
							"{0}{1:00000}@{2}", prifix, i, tb_emailAddress.Text.Split('@')[1]);
					emails.Add(email);
				}

				return string.Join(";", emails);
			}

			return i_partipient;
		}

		private ParallelLoopResult GenerateSendItems(string[] tos, string[] ccs, string[] bccs, Random random, int loopTime)
		{
			progressBar1.Maximum = loopTime * tos.Count();
			var index = 0;
			return Parallel.ForEach(tos, new ParallelOptions { MaxDegreeOfParallelism = threadCount }, (Action<string>)(to_sender =>
			 {
				 var service = Service;
				 service.Url = new Uri(m_urls[index % m_urls.Count]);
				 index++;
				 SetCredential(service, to_sender);
				 Random random2 = new Random();
				 var retryTimes = 0;
				 for (int i = 0; i < loopTime; i++)
				 {
					 try
					 {
						 EmailMessage message = new EmailMessage(service);
						 message.Sender = to_sender;
						 var recipient = "MS1User" + random2.Next(5000).ToString("00000") + "@365sale1.onmicrosoft.com";
						 message.ToRecipients.Add(recipient);
						 if (ccs[0] != string.Empty) message.CcRecipients.AddRange(ccs);
						 if (bccs[0] != string.Empty) message.ToRecipients.AddRange(bccs);
						 message.Subject = tb_subject.Text + to_sender + i.ToString("00000");

						 message.Body = tb_body.Text + to_sender + i.ToString("00000");
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

						 message.SendAndSaveCopy();
						 WriteLog(@".\nomal.txt", string.Format("{0}Sucessuce=>SenderEmail:{1}--Recipient:{2}--Loop:{3}\r\n", DateTime.Now, to_sender, recipient, i));
					 }
					 catch (Exception ex)
					 {
						 WriteLog(@".\error.txt", string.Format("{0}GenerateSendItems=>SenderEmail:{1}--Loop:{2}--Error:{3}\r\n", DateTime.Now, to_sender, i, ex.Message));
						 retryTimes++;
						 if (retryTimes < 3)
						 {
							 Thread.Sleep(TimeSpan.FromSeconds(10));
							 i--;
						 }
						 else
						 {
							 retryTimes = 0;

							 continue;
						 }

						 //File.AppendAllText(@".\error.txt", ex.Message + "\r\n");
					 }

					 progressBar1.Value++;
					//Thread.Sleep(TimeSpan.FromSeconds(5));
				}
			 }));
		}

		private static void WriteLog(string i_path, string i_content)
		{
			try
			{
				LogWriteLock.EnterWriteLock();
				File.AppendAllText(i_path, i_content);
				LogWriteLock.ExitWriteLock();
			}
			catch { }
		}

		private ParallelLoopResult SendByTos(string[] tos, string[] ccs, string[] bccs, Random random, int loopTime,bool sendAndSaveCopy)
		{
			var index = 0;
			return Parallel.ForEach(tos, new ParallelOptions { MaxDegreeOfParallelism = threadCount }, (Action<string>)(to_sender =>
			{
				var service = Service;
				service.Url = new Uri(m_urls[index % m_urls.Count]);
				index++;
				SetCredential(service, to_sender);

				foreach (var to in tos)
				{
					var retryTimes = 0;
					for (int i = 0; i < loopTime; i++)
					{
						try
						{
							EmailMessage message = new EmailMessage(service);

							if (to_sender.Contains("Share"))
							{
								message.From = new EmailAddress(to_sender);
							}
							else
							{
								message.Sender = to_sender;
							}
							message.ToRecipients.Add(to);
							if (ccs[0] != string.Empty) message.CcRecipients.AddRange(ccs);
							if (bccs[0] != string.Empty) message.ToRecipients.AddRange(bccs);
							message.Subject = tb_subject.Text + to_sender + i.ToString("00000");

							message.Body = tb_body.Text + to_sender + i.ToString("00000");
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
							if (sendAndSaveCopy)
							{
								if (to_sender.Contains("Share"))
								{
									Mailbox sharedMailbox = new Mailbox(to_sender);
									FolderId shareSenItems = new FolderId(WellKnownFolderName.SentItems, sharedMailbox);
									message.SendAndSaveCopy(shareSenItems);
								}
								else
								{
									message.SendAndSaveCopy();
								}
							}
							else
							{
								message.Send();
							}
						}
						catch(Exception ex)
						{


							WriteLog(@".\error.txt", string.Format("{0}SendByTos=>SenderEmail:{1}--Recipient:{2}--Loop:{3}--Error:{4}\r\n:", DateTime.Now, to_sender, to, i, ex.Message));

							retryTimes++;
							if(retryTimes < 3)
							{
								Thread.Sleep(TimeSpan.FromMinutes(3));
								i--;
							}
							else
							{
								retryTimes = 0;
								
								continue;
							}
						}

						//Thread.Sleep(TimeSpan.FromSeconds(5));
					}

					progressBar1.Value++;
				}
			}));
		}

		private ParallelLoopResult MoveInboxEmailToCustomFolder(
			String[] i_tos, string i_subFolder, int i_moveCount)
		{
			var index = 0;
			return Parallel.ForEach(i_tos, new ParallelOptions { MaxDegreeOfParallelism = threadCount }, (Action<string>)(to_sender =>
			{
				var service = Service;
				service.Url = new Uri(m_urls[index % m_urls.Count]);
				index++;
				SetCredential(service,to_sender);
				FolderId folderToAccess = new FolderId(WellKnownFolderName.Inbox, to_sender);

				Folder inboxfolder = to_sender.Contains("Share") ?
					Folder.Bind(service, folderToAccess) :
					Folder.Bind(service, WellKnownFolderName.Inbox);
				inboxfolder.Load();

				// Finds the emails in a certain folder, in this case the Junk Email
				FindItemsResults<Item> findResults = service.FindItems(inboxfolder.Id, new ItemView(i_moveCount));

				foreach (var folder in inboxfolder.FindFolders(new FolderView(100)))
				{
					// This IF limits what folder the program will seek
					if (folder.DisplayName == i_subFolder)
					{
						// Trust me, the ID is a pain if you want to manually copy and paste it. This stores it in a variable
						var fid = folder.Id;
						var items = findResults.Items;
						var retryTimes = 0;
						for (var i = 0; i < Math.Min(i_moveCount, items.Count); i++)
						{
							// Load the email, move the email into the id.  Note that MOVE needs a valid ID, which is why storing the ID in a variable works easily.
							
							try
							{
								items[i].Load();
								items[i].Move(fid);
							}
							catch (Exception ex)
							{
								WriteLog(@".\error.txt", string.Format("{0}MoveInboxEmailToCustomFolder=>SenderEmail:{1}--Loop:{2}--Error:{3}\r\n", DateTime.Now, to_sender, i, ex.Message));
								retryTimes++;
								if(retryTimes < 3)
								{
									Thread.Sleep(TimeSpan.FromMinutes(3));
									i--;
								}
								else
								{
									retryTimes = 0;
									continue;
								}
							}

						}
					}
				}
				progressBar1.Value++;
			}));
		}

		private ParallelLoopResult CreateFolder(string[] i_tos)
		{
			var i = 0;
			return Parallel.ForEach(i_tos, new ParallelOptions { MaxDegreeOfParallelism = threadCount }, (Action<string>)(to_sender =>
			{
				try
				{
					var service = Service;
					service.Url = new Uri(m_urls[i % m_urls.Count]);
					i++;
					SetCredential(service, to_sender);
					// Create a custom folder class.

					Folder folder = new Folder(service);
					folder.DisplayName = "Custom Folder";
					folder.FolderClass = "IPF.MyCustomFolderClass";

					FolderId folderToAccess = new FolderId(WellKnownFolderName.Inbox, to_sender);

					Folder inboxfolder = to_sender.Contains("Share") ?
						Folder.Bind(service, folderToAccess) :
						Folder.Bind(service, WellKnownFolderName.Inbox);

					var inboxSubFolders = inboxfolder.FindFolders(new FolderView(1000));


					//var desFolderCount = inboxSubFolders.Count(f => f.DisplayName == folder.DisplayName);
					if (inboxSubFolders.Count(f => f.DisplayName == folder.DisplayName) <= 0)
					{
						// Create the folder as a child of the Inbox folder.
						folder.Save(inboxfolder.Id);
					}
				}
				catch (Exception ex)
				{
					WriteLog(@".\error.txt", string.Format("{0}CreateFolder=>SenderEmail:{1}--Error:{2}\r\n", DateTime.Now, to_sender, ex.Message));

				}

				progressBar1.Value++;
			}));
		}

		private void SetCredential(ExchangeService i_service, string to_sender)
		{
			if (i_service.Url.ToString() == "https://outlook.office365.com/EWS/Exchange.asmx")
			{
				password = "I@mrootPWD";
			}
			if (to_sender.Contains("Share"))
			{
				i_service.Credentials = new WebCredentials(tb_emailAddress.Text, tb_password.Text);
			}
			else
			{
				i_service.Credentials = new WebCredentials(to_sender, password);
			}
		}

		private void bt_connect_Click(object sender, EventArgs e)
		{
			TrustCertificates();

			if (!Regex.IsMatch(tb_emailAddress.Text, EmailRegex))
			{
				MessageBox.Show(string.Format(
					"{0} has wrong email address format.", tb_emailAddress.Text));
			}

			Service.Credentials = new WebCredentials(tb_emailAddress.Text, tb_password.Text);

			m_urls.AddRange(tb_url.Text.Split(';'));
		}

		private static void TrustCertificates()
		{
			//Trust all certificates
			ServicePointManager.ServerCertificateValidationCallback =
				((sender1, certificate, chain, sslPolicyErrors) => true);

			
			ServicePointManager.ServerCertificateValidationCallback +=
				new RemoteCertificateValidationCallback(ValidateRemoteCertificate);
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
