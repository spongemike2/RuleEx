using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Serialization;
using Forms = System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

//=====================================================================================================================================================================================================
//=====================================================================================================================================================================================================
namespace RuleEx2025
{
	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	public abstract class Action
	{
		private static long id = 0;
		[XmlIgnore] public long Id;

		//=============================================================================================================================================================================================
		// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_members.aspx
		//=============================================================================================================================================================================================
		public abstract Outlook.MailItem Invoke(Rule parent, Outlook.MailItem mailItem);
		public abstract Outlook.MeetingItem Invoke(Rule parent, Outlook.MeetingItem mailItem);

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		protected void nop(object o=null)
		{
		}

		public Action()
		{
			this.Id = ++Action.id;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		protected void AddCategory(Outlook.MailItem mailItem, string category)
		{
			string oldCategories = mailItem.Categories;
			List<string> categoriesList = new List<string>();
			HashSet<string> categoriesSet = new HashSet<string>();

			if (oldCategories != null)
			{
				foreach(string c in oldCategories.Split(','))
				{
					categoriesList.Add(c);
					categoriesSet.Add(c);
				}
			}

			if (!categoriesSet.Contains(category))
			{
				categoriesSet.Add(category);
				categoriesList.Add(category);
			}

			categoriesList.Sort();

			string newCategories = string.Join(",", categoriesList.ToArray<string>());

			if (newCategories != oldCategories)
			{
				mailItem.Categories = newCategories;
				mailItem.Save();
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		protected void RemoveCategory(Outlook.MailItem mailItem, string category)
		{
			string oldCategories = mailItem.Categories;
			List<string> categoriesList = new List<string>();
			HashSet<string> categoriesSet = new HashSet<string>();

			if (oldCategories != null)
			{
				foreach(string c in oldCategories.Split(','))
				{
					categoriesList.Add(c);
					categoriesSet.Add(c);
				}
			}

			if (categoriesSet.Contains(category))
			{
				categoriesSet.Remove(category);
				categoriesList.Remove(category);
			}

			categoriesList.Sort();

			string newCategories = string.Join(",", categoriesList.ToArray<string>());

			if (newCategories != oldCategories)
			{
				mailItem.Categories = newCategories;
				mailItem.Save();
			}
		}
	}

	//=================================================================================================================================================================================================
	// The build move action is a special case of move: It moves the email into a *sub* folder of the build based on the build recipeint
	//=================================================================================================================================================================================================
	public class BuildMoveAction : Action
	{
		[XmlAttribute] public string ParentFolderName;
		[XmlAttribute] public string ParentFolderPath;
		[XmlElement("BuildEmail")] public string[] Recipients;
		[XmlElement("BuildAccount")] public string[] BuildAccounts;

		[XmlIgnore] public Outlook.MAPIFolder ParentFolder;

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public BuildMoveAction() : base()
		{
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public BuildMoveAction(string parentFolderName, string[] recipients, string[] buildAccounts) : base()
		{
			this.ParentFolderName	= parentFolderName;
			this.Recipients			= recipients;
			this.BuildAccounts		= buildAccounts;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private object _Invoke(dynamic mailItem)
		{
			// get the parent folder
			if (this.ParentFolder == null)
			{
				var application = mailItem.Application;
				this.ParentFolder = application.Session.GetFolderFromID(this.ParentFolderName);
				this.ParentFolderPath = this.ParentFolder.FullFolderPath;
			}

			if (this.ParentFolder != null)
			{
				Outlook.Recipient thisRecipient = null;

				foreach (Outlook.Recipient mailRecipient in mailItem.Recipients)
				{
					foreach (string actionRecipient in this.Recipients)
					{
						string s1 = mailRecipient.Address.ToLower();
						string s2 = actionRecipient.ToLower();

						if (s1 == s2)
						{
							thisRecipient = mailRecipient;
							break;
						}
					}

					if (thisRecipient != null)
					{
						break;
					}
				}

				if (thisRecipient != null)
				{
					string childFolderName = thisRecipient.Name;
					// if the action recipient is an email address, just consider the name part
					var atSignIndex = childFolderName.IndexOf("@");

					if (atSignIndex != -1)
					{
						childFolderName = childFolderName.Substring(0, atSignIndex);
					}

					// remove "Build Info", if we have it
					childFolderName = Regex.Replace(childFolderName, "^(.*?) ?Build Info(.*)$", "$1$2");

					// okay, now we have the recipient
					Regex regex = new Regex(string.Format(@"(({0}\s)|({0}$))", Regex.Escape(childFolderName)), RegexOptions.IgnoreCase);
					Outlook.MAPIFolder childFolder = this.ParentFolder.Folders.Cast<Outlook.MAPIFolder>().Where(f => regex.IsMatch(f.Name)).FirstOrDefault();

					if (childFolder == null)
					{
						// create it
						childFolder = this.ParentFolder.Folders.Add(childFolderName.ToLower());
						this.nop();
					}
					else
					{
						// use it
						this.nop();
					}

					Logger.the.WriteLineFormat("   Recipient: {0}", thisRecipient.Name);
					Logger.the.WriteLineFormat("   Folder: {0}", childFolder.FullFolderPath);

					if (mailItem.Parent.FullFolderPath != childFolder.FullFolderPath)
					{
						try
						{
							mailItem = mailItem.Move(childFolder);
						}
						catch
						{
							// it didn't work, log it, and move on
							// $ToDo: log it
						}
					}

					bool sentFromBuildAccount = false;

					foreach (string buildAccount in this.BuildAccounts)
					{
						string sender;

						if (mailItem is Outlook.MailItem)
						{
							sender = mailItem.Sender.Address;
						}
						else if (mailItem is Outlook.MeetingItem)
						{
							sender = mailItem.SenderEmailAddress;
						}
						else
						{
							sender = string.Empty;
						}

						if (sender == buildAccount)
						{
							sentFromBuildAccount = true;
							break;
						}
					}

					bool markAsRead = false;
					bool markBuildMonitorEmailsAsRead = true;
					//RegexOptions options = RegexOptions.IgnoreCase;
					RegexOptions options = RegexOptions.None;

					if (sentFromBuildAccount)
					{
						// okay, since it's sent from a build account, we will mark it as read unless it's flagged as important or
						// has "errors" in the subject
						if ((mailItem is Outlook.MailItem) && (mailItem.Importance == (int)Outlook.OlImportance.olImportanceHigh))
						{
							if (false) {}
							else if (markBuildMonitorEmailsAsRead && (Regex.IsMatch(mailItem.Subject, "^Build Monitor found errors:", options)))
							{
								markAsRead = true;
							}
							else if (Regex.IsMatch(mailItem.Subject, "error|failed", RegexOptions.IgnoreCase))
							{
								this.AddCategory(mailItem, "BuildError");
								if (Regex.IsMatch(mailItem.Subject, " BCR$", RegexOptions.IgnoreCase))
								{
									markAsRead = true;
								}
								else
								{
									markAsRead = false;
								}
							}
							else if (Regex.IsMatch(mailItem.Subject, "warnings", RegexOptions.IgnoreCase))
							{
								markAsRead = true;
							}
							else
							{
								markAsRead = false;
							}
						}
						else
						{
							if (Regex.IsMatch(mailItem.Subject, "error|fail", RegexOptions.IgnoreCase))
							{
								markAsRead = false;
							}
							else
							{
								markAsRead = true;
							}
						}

					}
					else
					{
						// *DON'T* mark it as read. We're done here.
						markAsRead = false;
					}

					if (markAsRead)
					{
						mailItem.UnRead = false;
					}
				}


				// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mapifolder_members.aspx


				// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_members.aspx

				// find the folder where this email goes...

			}

			return mailItem;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public override Outlook.MailItem Invoke(Rule parent, Outlook.MailItem mailItem)
		{
			return this._Invoke(mailItem) as Outlook.MailItem;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public override Outlook.MeetingItem Invoke(Rule parent, Outlook.MeetingItem mailItem)
		{
			return this._Invoke(mailItem) as Outlook.MeetingItem;
		}
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	public class MoveAction : Action
	{
		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		[XmlAttribute] public string FolderName;
		[XmlAttribute] public string FolderPath;
		[XmlIgnore] public Outlook.MAPIFolder Folder;

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public MoveAction() : base()
		{
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public MoveAction(string folder) : base()
		{
			this.FolderName = folder;
		}

		//=============================================================================================================================================================================================
		// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_members.aspx
		//=============================================================================================================================================================================================
		private object _Invoke(Rule parent, dynamic mailItem)
		{
			int retryCount = 0;
			int maxRetries = 1;

			if (this.Folder == null)
			{
				var application = mailItem.Application;

				while (true)
				{
					++retryCount;

					try
					{
						this.Folder = application.Session.GetFolderFromID(this.FolderName);
						this.FolderPath = this.Folder.FullFolderPath;
						break;
					}
					catch (Exception ex)
					{
						Logger.the.WriteFormat("Error processing rule ({0}): folder does not exist. Exception: {1}", this.FolderName, ex.Message);

						if (retryCount > maxRetries)
						{
							var response = Forms.MessageBox.Show(string.Format("Rule {0} has an invalid folder in the Move action. Would you like to pick a new folder?", parent.Name), "Invalid folder", Forms.MessageBoxButtons.YesNo);
							if (response == Forms.DialogResult.Yes)
							{
								var folder = ThisAddIn.theApplication.Session.PickFolder();
								this.nop(folder);

								if (folder != null)
								{
									this.Folder = folder;
									this.FolderName = folder.EntryID;
									this.FolderPath = folder.FolderPath;
									return this._Invoke(parent, mailItem);
								}
							}

							return mailItem;
						}
					}
				}
			}

			if (this.Folder != null)
			{
				var mailFolder = mailItem.Parent as Outlook.MAPIFolder;
				bool loop = true;

				while (loop)
				{
					++retryCount;

					try
					{
						if (mailFolder != null)
						{
							if (mailFolder.FullFolderPath != this.Folder.FullFolderPath)
							{
								return mailItem.Move(this.Folder);
							}
						}

						loop = false;
					}
					catch (Exception ex)
					{
						Logger.the.WriteFormat("Error processing rule ({0}): folder does not exist. Exception: {1}", this.FolderName, ex.Message);

						if (retryCount > maxRetries)
						{
							var response = Forms.MessageBox.Show(string.Format("Rule {0} has an invalid folder in the Move action. Would you like to pick a new folder?", parent.Name), "Invalid folder", Forms.MessageBoxButtons.YesNo);
							if (response == Forms.DialogResult.Yes)
							{
								var folder = ThisAddIn.theApplication.Session.PickFolder();
								this.nop(folder);

								if (folder != null)
								{
									this.Folder = folder;
									this.FolderName = folder.EntryID;
									this.FolderPath = folder.FolderPath;
									return this._Invoke(parent, mailItem);
								}
							}

							return mailItem;
						}
					}
				}
			}

			return mailItem;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public override Outlook.MailItem Invoke(Rule parent, Outlook.MailItem mailItem)
		{
			return this._Invoke(parent, mailItem) as Outlook.MailItem;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public override Outlook.MeetingItem Invoke(Rule parent, Outlook.MeetingItem mailItem)
		{
			return this._Invoke(parent, mailItem) as Outlook.MeetingItem;
		}
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	public class MarkAsReadAction : Action
	{
		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public MarkAsReadAction() : base()
		{
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public override Outlook.MailItem Invoke(Rule parent, Outlook.MailItem mailItem)
		{
			mailItem.UnRead = false;
			return mailItem;
		}
		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public override Outlook.MeetingItem Invoke(Rule parent, Outlook.MeetingItem mailItem)
		{
			mailItem.UnRead = false;
			return mailItem;
		}
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	public class TagCategoryAction : Action
	{
		public enum OperationType
		{
			Add,
			Remove,
		}

		[XmlAttribute] public OperationType Operation;
		[XmlAttribute] public string Category;


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public TagCategoryAction() : base()
		{
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public TagCategoryAction(string category, OperationType operation) : base()
		{
			this.Category = category;
			this.Operation = operation;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public override Outlook.MailItem Invoke(Rule parent, Outlook.MailItem mailItem)
		{
			if (this.Operation == OperationType.Add)
			{
				AddCategory(mailItem, this.Category);
			}
			else
			{
				RemoveCategory(mailItem, this.Category);
			}

			return mailItem;
		}
		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public override Outlook.MeetingItem Invoke(Rule parent, Outlook.MeetingItem mailItem)
		{
			return mailItem;
		}
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	public class ForwardAction : Action
	{
		[XmlAttribute] public string EmailAddress;

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public ForwardAction() : base()
		{
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public ForwardAction(string emailAddress) : base()
		{
			this.EmailAddress = emailAddress;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public override Outlook.MailItem Invoke(Rule parent, Outlook.MailItem mailItem)
		{
			if (this.EmailAddress.Length > 0)
			{
				try
				{
					var forwardItem = mailItem.Forward();
					forwardItem.Recipients.Add(this.EmailAddress);
					forwardItem.DeleteAfterSubmit = true;
					try
					{
						forwardItem.Send();
					}
					catch (Exception ex)
					{
						this.nop(ex);
					}
				}
				catch (Exception ex)
				{
					this.nop(ex);
				}
			}

			return mailItem;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public override Outlook.MeetingItem Invoke(Rule parent, Outlook.MeetingItem mailItem)
		{
			return mailItem;
		}
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	public class FixLinksAction : Action
	{
		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private const string regexStringHtml = @"(?<pre><a\s.*href="")(?<link>[^\""<>\s]*\.safelinks\.protection\.outlook\.com[^\""]*)(?<post>"")";
		private const string regexStringText = @"(?<pre>)(?<link>https?:[^\s]*\.safelinks\.protection\.outlook\.com[^\s]*)(?<post>)";
		private static Regex regexHtml = new Regex(regexStringHtml, RegexOptions.IgnoreCase | RegexOptions.Compiled);
		private static Regex regexText = new Regex(regexStringText, RegexOptions.IgnoreCase | RegexOptions.Compiled);


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public FixLinksAction() : base()
		{
		}


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private static string FixTextLinks(string text)
		{
			var newBody = new StringBuilder(text.Length);
			bool textDirty = false;

			int iteration = 0;
			do
			{
				++iteration;
				Logger.the.WriteLineFormat(@"Text loop iteration: {0}", iteration);

				newBody.Clear();

				var match = regexText.Match(text);

				if (match.Success)
				{
					textDirty = true;

					// now, fix the link!
					var pre = match.Groups["pre"].Value;
					var link = match.Groups["link"].Value;
					var post = match.Groups["post"].Value;

					var uri = new Uri(link);
					var query = uri.Query;
					var querystring = System.Web.HttpUtility.ParseQueryString(query);
					var newurl = querystring["url"];

					if ((newurl != null) && (newurl.Length > 0))
					{
						link = newurl;
					}
					else
					{
						break;
					}

					newBody.Append(text.Substring(0, match.Index));

					newBody.Append(pre);
					newBody.Append(link);
					newBody.Append(post);

					newBody.Append(text.Substring(match.Index + match.Length, text.Length - (match.Index + match.Length)));

					text = newBody.ToString();
				}
				else
				{
					break;
				}

			} while (true);


			if (textDirty)
			{
				return text;
			}

			return null;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private static string FixHtmlLinks(string html)
		{
			var newBody = new StringBuilder(html.Length);
			bool htmlDirty = false;

			int iteration = 0;
			do
			{
				++iteration;
				Logger.the.WriteLineFormat(@"Html loop iteration: {0}", iteration);

				newBody.Clear();

				var match = regexHtml.Match(html);

				if (match.Success)
				{
					htmlDirty = true;

					// now, fix the link!
					var pre = match.Groups["pre"].Value;
					var link = match.Groups["link"].Value;
					var post = match.Groups["post"].Value;

					var uri = new Uri(link);
					var query = uri.Query;
					var querystring = System.Web.HttpUtility.ParseQueryString(query);
					var newurl = querystring["url"];

					if ((newurl != null) && (newurl.Length > 0))
					{
						link = newurl;
					}
					else
					{
						break;
					}

					newBody.Append(html.Substring(0, match.Index));

					newBody.Append(pre);
					newBody.Append(link);
					newBody.Append(post);

					newBody.Append(html.Substring(match.Index + match.Length, html.Length - (match.Index + match.Length)));

					html = newBody.ToString();
				}
				else
				{
					break;
				}

			} while (true);


			if (htmlDirty)
			{
				return html;
			}

			return null;
		}

		//=============================================================================================================================================================================================
		// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_members.aspx
		//=============================================================================================================================================================================================
		public override Outlook.MailItem Invoke(Rule parent, Outlook.MailItem mailItem)
		{
			bool dirty = false;

			if (mailItem.HTMLBody != null)
			{
				var newHtmlBody = FixHtmlLinks(mailItem.HTMLBody);
				if (newHtmlBody != null)
				{
					mailItem.HTMLBody = newHtmlBody;
					dirty = true;
				}
			}

			if (mailItem.Body != null)
			{
				var newTextBody = FixTextLinks(mailItem.Body);
				if (newTextBody != null)
				{
					mailItem.Body = newTextBody;
					dirty = true;
				}
			}

			if (dirty)
			{
				mailItem.Save();
			}

			return mailItem;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public override Outlook.MeetingItem Invoke(Rule parent, Outlook.MeetingItem mailItem)
		{
			return mailItem;
		}
	}
}
