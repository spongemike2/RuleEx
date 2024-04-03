using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Serialization;
using Outlook = Microsoft.Office.Interop.Outlook;

//=====================================================================================================================================================================================================
//=====================================================================================================================================================================================================
namespace RuleEx2025
{
	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	public abstract class Condition
	{
		private static long id = 0;

		// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_members.aspx
		//public abstract bool Test(Outlook.MailItem mailItem);
		//public abstract bool Test(Outlook.MeetingItem mailItem);
		[XmlAttribute] public bool Not;
		[XmlIgnore] public long Id;


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public Condition Clone()
		{
			XmlSerializer ser = new XmlSerializer(this.GetType());
			MemoryStream stream = new MemoryStream();
			ser.Serialize(stream, this, new XmlSerializerNamespaces(new XmlQualifiedName[]{new XmlQualifiedName("")}));
			stream.Position = 0;
			Condition newCondition = ser.Deserialize(stream) as Condition;
			stream.Close();
			return newCondition;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		protected Condition()
		{
			this.Id = ++Condition.id;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		protected abstract bool _Test(dynamic mailItem, Dictionary<string,string> meta);

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public virtual bool Test(Outlook.MailItem mailItem, Dictionary<string,string> meta)
		{
			return this.Not ^ this._Test(mailItem, meta);
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public virtual bool Test(Outlook.MeetingItem mailItem, Dictionary<string,string> meta)
		{
			return this.Not ^ this._Test(mailItem, meta);
		}
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	public class RecipientCondition : Condition
	{
		[XmlAttribute] public string Recipient;
		[XmlAttribute] public string Regex;
		[XmlIgnore] private Regex _regex;

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private static void nop(object o = null)
		{
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public RecipientCondition() : base()
		{
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public RecipientCondition(string recipient, string regex) : base()
		{
			this.Recipient = recipient == null ? null : recipient.ToLower();
			this.Regex = regex;

			if (this._regex == null)
			{
				if (!string.IsNullOrEmpty(this.Regex))
				{
					this._regex = new Regex(this.Regex);
				}
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		protected override bool _Test(dynamic someItem, Dictionary<string,string> meta)
		{
			if (this._regex == null)
			{
				if (!string.IsNullOrEmpty(this.Regex))
				{
					this._regex = new Regex(this.Regex);
				}
			}

			if (this._regex != null)
			{
				nop();
			}


			if (someItem is Outlook.MailItem)
			{
				Outlook.MailItem mailItem = someItem as Outlook.MailItem;

				try
				{
					var recipients = mailItem.Recipients;
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.recipients_members.aspx
					foreach (Outlook.Recipient recipient in recipients)
					{
						// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.recipient_members.aspx

						if (recipient.Address != null)
						{
							string s1 = recipient.Address.ToLower();

							if (this._regex == null)
							{
								if (!string.IsNullOrEmpty(this.Recipient))
								{
									string s2 = this.Recipient.ToLower();

									if (s1 == s2)
									{
										return true;
									}
								}
							}
							else
							{
								// regex
								return this._regex.IsMatch(s1);
							}
						}
					}
				}
				catch (System.Exception ex)
				{
					System.Diagnostics.Debug.WriteLine("Exception: {0}", ex.Message);
					nop(ex);
				}
			}
			else if (someItem is Outlook.MeetingItem)
			{
				Outlook.MeetingItem mailItem = someItem as Outlook.MeetingItem;

				try
				{
					var recipients = mailItem.Recipients;
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.recipients_members.aspx
					foreach (Outlook.Recipient recipient in recipients)
					{
						// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.recipient_members.aspx
						if (recipient.Address != null)
						{
							string s1 = recipient.Address.ToLower();

							if (this._regex == null)
							{
								if (this.Recipient != null)
								{
									string s2 = this.Recipient.ToLower();

									if (s1 == s2)
									{
										return true;
									}
								}
							}
							else
							{
								// regex
								return this._regex.IsMatch(s1);
							}
						}
					}
				}
				catch (System.Exception ex)
				{
					System.Diagnostics.Debug.WriteLine("Exception: {0}", ex.Message);
					nop(ex);
				}

			}

			return false;
		}
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	public class SenderCondition : Condition
	{
		[XmlAttribute] public string Sender;
		[XmlAttribute] public string Regex;
		[XmlIgnore] private Regex _regex;

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public SenderCondition() : base()
		{
			if (!string.IsNullOrEmpty(this.Sender))
			{
				this.Sender = this.Sender.ToLower();
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public SenderCondition(string sender, string senderRegEx) : base()
		{
			this.Sender = sender==null ? null : sender.ToLower();
			this.Regex = senderRegEx;

			if (this._regex == null)
			{
				if (!string.IsNullOrEmpty(this.Regex))
				{
					this._regex = new Regex(this.Regex);
				}
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private void nop(object o=null)
		{
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		protected override bool _Test(dynamic item, Dictionary<string,string> meta)
		{
			if (this._regex == null)
			{
				if (!string.IsNullOrEmpty(this.Regex))
				{
					this._regex = new Regex(this.Regex);
				}
			}

			string s1 = "";
			string s2 = this.Sender==null ? "" : this.Sender.ToLower();

			if (meta.ContainsKey("Sender"))
			{
				s1 = meta["Sender"];
			}
			else
			{
				try
				{
					Outlook.MailItem mailItem = item as Outlook.MailItem;
					if (mailItem != null)
					{
						if (mailItem.Sender != null)
						{
							if (mailItem.Sender.Address != null)
							{
								var dt0 = DateTime.Now;
								s1 = mailItem.Sender.Address.ToLower();
								var dt1 = DateTime.Now;
								Logger.the.WriteLineFormat("It took {0} seconds to evaluate the sender \"{1}\".", (dt1-dt0).TotalSeconds, s1);
							}
						}
					}
					else
					{
						Outlook.MeetingItem meetingItem = item as Outlook.MeetingItem;

						if (meetingItem != null)
						{
							if (meetingItem.SenderName != null)
							{
								if (meetingItem.SenderEmailAddress != null)
								{
									s1 = meetingItem.SenderEmailAddress.ToLower();
								}
							}
						}
					}

					meta.Add("Sender", s1);
				}
				catch(Exception ex)
				{
					this.nop(ex);
				}
			}

			if (string.IsNullOrEmpty(s2))
			{
				return this._regex.IsMatch(s1);
			}
			else
			{
				return (s1 == s2);
			}
		}
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	public class SubjectCondition : Condition
	{
		[XmlAttribute] public string Regex;
		[XmlIgnore] private Regex _regex;

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public SubjectCondition() : base()
		{
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public SubjectCondition(string regex) : base()
		{
			this.Regex = regex;

			if (this._regex == null)
			{
				if (!string.IsNullOrEmpty(this.Regex))
				{
					this._regex = new Regex(this.Regex);
				}
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_members.aspx
		protected override bool _Test(dynamic mailItem, Dictionary<string,string> meta)
		{
			if (this._regex == null)
			{
				if (!string.IsNullOrEmpty(this.Regex))
				{
					this._regex = new Regex(this.Regex);
				}
			}

			return this._regex.IsMatch(mailItem.Subject);
		}
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	public class BodyCondition : Condition
	{
		[XmlAttribute] public string Regex;
		[XmlIgnore] private Regex _regex;

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public BodyCondition() : base()
		{
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public BodyCondition(string regex) : base()
		{
			this.Regex = regex;

			if (this._regex == null)
			{
				if (!string.IsNullOrEmpty(this.Regex))
				{
					this._regex = new Regex(this.Regex);
				}
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_members.aspx
		protected override bool _Test(dynamic mailItem, Dictionary<string,string> meta)
		{
			if (this._regex == null)
			{
				if (!string.IsNullOrEmpty(this.Regex))
				{
					this._regex = new Regex(this.Regex);
				}
			}

			string body = mailItem.Body;
			bool result = this._regex.IsMatch(body);
			return result;
		}
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	public class ConditionGroup : Condition
	{
		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public enum GroupingOperator
		{
			And,
			Or,
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		[XmlElement(typeof(RecipientCondition))]
		[XmlElement(typeof(SenderCondition))]
		[XmlElement(typeof(SubjectCondition))]
		[XmlElement(typeof(BodyCondition))]
		[XmlElement(typeof(ConditionGroup))]
		public Condition[]	Conditions;
		[XmlAttribute] public GroupingOperator Operator;

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public ConditionGroup() : base()
		{
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public ConditionGroup(GroupingOperator groupingOperator, Condition[] conditions) : base()
		{
			this.Operator = groupingOperator;
			this.Conditions = conditions;
		}

		//=============================================================================================================================================================================================
		// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_members.aspx
		//=============================================================================================================================================================================================
		protected override bool _Test(dynamic mailItem, Dictionary<string,string> meta)
		{
			if (this.Operator == GroupingOperator.And)
			{
				foreach (Condition condition in this.Conditions)
				{
					if (!condition.Test(mailItem, meta))
					{
						return false;
					}
				}

				return true;
			}
			else
			{
				foreach (Condition condition in this.Conditions)
				{
					if (condition.Test(mailItem, meta))
					{
						return true;
					}
				}

				return false;
			}
		}
	}
}
