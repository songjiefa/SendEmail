using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SentEmail
{
	public class EwsInfo
	{
		public String URL { get; set; }
		public String LoginEmail { get; set; }
		public String Password { get; set; }
		public String To;
		public String Cc;
		public String Bcc;
		public String Subject;
		public String Body;
		public String Attachments;
		public Int32 LoopTimes;
		public float AttachmentRate;
	}
}
