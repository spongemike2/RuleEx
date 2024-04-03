using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

//=====================================================================================================================================================================================================
//=====================================================================================================================================================================================================
namespace RuleEx2025
{
	//=================================================================================================================================================================================================
	// https://msdn.microsoft.com/en-us/library/microsoft.office.tools.outlook.outlookaddin.aspx
	//=================================================================================================================================================================================================
	public partial class ThisAddIn
	{
		//
		// needed in order to ensure that the "OnMailArrived" event is maintained
		//
		Outlook.Items _inboxItems;

		private ManualResetEvent _quit;
		private Settings _settings;
		private string _settingsFile;
		private Logger _logger;
		private Mutex _runMutex = new Mutex();
		private bool _cancelRun = false;

		public static Outlook.Application theApplication;

		//
		// Patterns used to convert from a DL's display name to the name of the folder
		//
		private string[] patterns = new string[] {
			@"^(?<pre>Michael J. Lyons \()(?<alias>.*)(?<post>\))$",
			@"^(?<pre>Mike Lyons \()(?<alias>.*)(?<post>\))$",
			@"^(?<pre>Mike and Vanessa Lyons \()(?<alias>.*)(?<post>\))$",
			@"^(?<pre>Michael and Vanessa Lyons \()(?<alias>.*)(?<post>\))$",
			@"^(?<pre>Vanessa Lyons \()(?<alias>.*)(?<post>\))$",
			@"^(?<pre>Vanessa L\. Lyons \()(?<alias>.*)(?<post>\))$",
			@"^(?<pre>Megan Lyons \()(?<alias>.*)(?<post>\))$",
			@"^(?<pre>Ryne Lyons \()(?<alias>.*)(?<post>\))$",
		};

		//
		// Folder names for finding an existing folder
		//
		private string[] foldernames = new string[] {
			@"Michael J. Lyons ({0})",
			@"Mike Lyons ({0})",
			@"Mike and Vanessa Lyons ({0})",
			@"Michael and Vanessa Lyons ({0})",
			@"Vanessa Lyons ({0})",
			@"Vanessa L. Lyons ({0})",
			@"Megan Lyons ({0})",
			@"Ryne Lyons ({0})",
		};


		//=============================================================================================================================================================================================
		// useful for debugging
		//=============================================================================================================================================================================================
		private void nop(object o=null)
		{
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private Outlook.ApplicationEvents_11_Event _appEvents;

		private void ThisAddIn_Startup(object sender, System.EventArgs e)
		{
			if (!this.StartExclusive())
			{
				return;
			}

			this._logger = new FileLogger();
			ThisAddIn.theApplication = this.Application;

			// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.application_members.aspx
			// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.namespace_members.aspx
			// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mapifolder_members.aspx
			// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.items_members.aspx
			// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.itemsevents_event.itemadd.aspx
			// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.itemsevents_itemaddeventhandler.aspx
			// http://blogs.msdn.com/b/vsto/archive/2009/12/15/making-a-custom-group-appear-in-the-message-tab-of-a-mail-item-norm-estabrook.aspx
			// http://blogs.msdn.com/b/vsto/archive/2009/12/15/making-a-custom-group-appear-in-the-message-tab-of-a-mail-item-norm-estabrook.aspx
			// Application:
			//    https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.application
			//
			// Explorer:
			//    https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.explorer
			//
			// CommandBar:
			//    https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.core.commandbars
			//
			//
			// https://msdn.microsoft.com/en-us/library/office/microsoft.office.core.commandbars_members.aspx
			// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.application_members.aspx
			// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.namespace_members.aspx

			// http://www.kebabshopblues.co.uk/2007/01/04/visual-studio-2005-tools-for-office-commandbarbutton-faceid-property/
			// http://www.kebabshopblues.co.uk/2007/01/24/outlook-faceid-3000-3999/

			this._appEvents = this.Application as Outlook.ApplicationEvents_11_Event;
			if (this._appEvents == null)
			{
				Logger.the.WriteLine("Error, could not get application Outlook.ApplicationEvents_11_Event object");
				return;
			}

			//
			// create the quit event
			//
			this._quit = new ManualResetEvent(false);

			//
			// Register to be notified if we quit
			//
			this._appEvents.Quit += new Outlook.ApplicationEvents_11_QuitEventHandler(ThisAddIn_OnQuit);

			//
			// set the location of the settings file
			//
			this._settingsFile = Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments), "RuleExConfig.xml");

			//
			// load the settings
			//
			try
			{
				this._settings = Settings.Load(this._settingsFile);
			}
			catch (Exception ex)
			{
				this.nop(ex);
				this._settings = new Settings();
				this._settings.Resolve(this.Application);
				this._settings.Save(this._settingsFile);
			}

			Outlook.Explorers selectExplorers = this.Application.Explorers;
			selectExplorers.NewExplorer +=new Outlook.ExplorersEvents_NewExplorerEventHandler(newExplorer_Event);

			if (this._settings.Active)
			{
				this._inboxItems = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Items;
				this._inboxItems.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(OnMailArrived);
				this._inboxItems.ItemChange += new Outlook.ItemsEvents_ItemChangeEventHandler(OnMailChanged);
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private void newExplorer_Event(Outlook.Explorer new_Explorer)
		{
			((Outlook._Explorer)new_Explorer).Activate();
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private void OnMailArrived(object item)
		{
			// do stuff
			if (item is Outlook.MailItem)
			{
				// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_members.aspx
				var mailItem = item as Outlook.MailItem;
				this.RunRulesOnMailItem(mailItem);
			}
			else if (item is Outlook.MeetingItem)
			{
				// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.meetingitem_members.aspx
				var mailItem = item as Outlook.MeetingItem;
				this.RunRulesOnMailItem(mailItem);
			}
			else
			{
				this.nop();
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private void OnMailChanged(object item)
		{
			// do stuff
			if (item is Outlook.MailItem)
			{
				// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_members.aspx
				var mailItem = item as Outlook.MailItem;
			}
			else
			{
				this.nop();
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private void ThisAddIn_OnQuit()
		{
			this._quit.Set();
			this._settings.Save(this._settingsFile);
			this.nop();
			this.EndExclusive();
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
		{
			// Note: Outlook no longer raises this event. If you have code that
			//    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
		}

		#region UserInterface

		private void Inspectors_NewInspector(Outlook.Inspector Inspector)
		{
			throw new NotImplementedException();
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private bool SearchDistributionListRecursivelyForUser(Outlook.ExchangeDistributionList distributionList, string searchFor, ref bool cancel, ProgressDialog dialog=null, ArrayList list=null, int depth=1, HashSet<string> alreadyScannedSet=null)
		{
			if (cancel)
			{
				return false;
			}

			if (list == null)
			{
				list = new ArrayList();
			}

			if (alreadyScannedSet == null)
			{
				alreadyScannedSet = new HashSet<string>();
			}

			// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.exchangedistributionlist_members.aspx
			//Outlook.AddressEntries members = distributionList.GetMemberOfList();
			try
			{
				DateTime t0 = DateTime.Now;
				Outlook.AddressEntries members = null;
				members = distributionList.GetExchangeDistributionListMembers();

				if (dialog != null)
				{
					dialog.TextBoxText = string.Format("Processing DL: {0}", distributionList.Name);
				}

				foreach (Outlook.AddressEntry member in members)
				{
					//string padding = new string(' ', 4 * depth);
					//string output = string.Format("Scanning User: {1}", padding, member.Name);
					//this._logger.WriteLine(output);

					string memberAddress = member.Address;
					var memberType = member.DisplayType;

					if (member.Address == searchFor)
					{
						list.Add(member.Name);
						return true;
					}

					if (alreadyScannedSet.Contains(member.Address))
					{
						continue;
					}

					alreadyScannedSet.Add(member.Address);

					if (member.DisplayType == Outlook.OlDisplayType.olDistList)
					{
#if DEBUG
						string name = member.Name;
						string type = member.Type;
						string classname = member.Class.ToString();
						string displayType = member.DisplayType.ToString();

						this._logger.WriteLineFormat("{0} {1} {2} {3}", name, type, classname, displayType);
#endif

						Outlook.ExchangeDistributionList subDistributionList = member.GetExchangeDistributionList();

						if (subDistributionList != null)
						{
							if (SearchDistributionListRecursivelyForUser(subDistributionList, searchFor, ref cancel, dialog, list, depth + 1, alreadyScannedSet))
							{
								list.Add(member.Name);
								return true;
							}
						}
					}
				}
			}
			catch (Exception ex)
			{
				this.nop(ex);
				return false;
			}

			return false;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private string[] DetermineWhyMe(string searchFor, Outlook.MailItem mailItem, ref bool cancel, ProgressDialog dialog=null)
		{
			ArrayList list = new ArrayList();

			foreach (Outlook.Recipient recipient in mailItem.Recipients)
			{
				list.Clear();

				string recipientAddress = recipient.Address;

				if (recipientAddress.ToLower() == searchFor.ToLower())
				{
					// $ToDo
					list.Add(searchFor);
					break;
				}
				else
				{
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.recipient_members.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.addressentry_members.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.exchangedistributionlist_members.aspx
					Outlook.ExchangeDistributionList distributionList = recipient.AddressEntry.GetExchangeDistributionList();
					if (dialog != null)
					{
						dialog.TextBoxText = string.Format("Expanding DL: {0}", recipient.Name);
					}

					this._logger.WriteLine(recipient.Name);

					if (distributionList != null)
					{
						if (SearchDistributionListRecursivelyForUser(distributionList, searchFor, ref cancel, dialog, list))
						{
							list.Add(recipient.Name);
							list.Add(recipient.Address);
							break;
						}
					}
				}
			}

			return list.ToArray(typeof(string)) as string[];
		}

		private string[] DetermineWhyMe(string searchFor, Outlook.MeetingItem mailItem, ref bool cancel, ProgressDialog dialog=null)
		{
			ArrayList list = new ArrayList();

			foreach (Outlook.Recipient recipient in mailItem.Recipients)
			{
				list.Clear();

				string recipientAddress = recipient.Address;

				if (recipientAddress.ToLower() == searchFor.ToLower())
				{
					// $ToDo
					list.Add(searchFor);
					break;
				}
				else
				{
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.recipient_members.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.addressentry_members.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.exchangedistributionlist_members.aspx
					Outlook.ExchangeDistributionList distributionList = recipient.AddressEntry.GetExchangeDistributionList();
					if (dialog != null)
					{
						dialog.TextBoxText = string.Format("Expanding DL: {0}", recipient.Name);
					}

					this._logger.WriteLine(recipient.Name);

					if (distributionList != null)
					{
						if (SearchDistributionListRecursivelyForUser(distributionList, searchFor, ref cancel, dialog, list))
						{
							list.Add(recipient.Name);
							list.Add(recipient.Address);
							break;
						}
					}
				}
			}

			return list.ToArray(typeof(string)) as string[];
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private string[] DetermineWhyMe1(string searchFor, Outlook.MailItem mailItem, ref bool cancel, ProgressDialog dialog=null)
		{
			ArrayList list = new ArrayList();

			foreach (Outlook.Recipient recipient in mailItem.Recipients)
			{
				list.Clear();

				string recipientAddress = recipient.Address;

				if (recipientAddress.ToLower() == searchFor.ToLower())
				{
					// $ToDo
					list.Add(searchFor);
					break;
				}
				else
				{
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.recipient_members.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.addressentry_members.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.exchangedistributionlist_members.aspx
					Outlook.ExchangeDistributionList distributionList = recipient.AddressEntry.GetExchangeDistributionList();
					if (dialog != null)
					{
						dialog.TextBoxText = string.Format("Expanding DL: {0}", recipient.Name);
					}

					this._logger.WriteLine(recipient.Name);

					if (distributionList != null)
					{
						if (SearchDistributionListRecursivelyForUser(distributionList, searchFor, ref cancel, dialog, list))
						{
							list.Add(recipient.Name);
							list.Add(recipient.Address);
							break;
						}
					}
				}
			}

			return list.ToArray(typeof(string)) as string[];
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private bool DetermineWhyMeWithDialog(dynamic mailItem, ArrayList hierarchy)
		{
			bool cancel = false;
			bool itemProcessed = false;
			string[] whyMe = new string[0];

			// create the uber "WhyMe" task that will run in the NON-UI thread
			var WhyMeTask = System.Threading.Tasks.Task.Run(() =>
			{
				ProgressDialog dialog = new ProgressDialog();

				// create the progress bar task that merely mimics time
				var progressTask = System.Threading.Tasks.Task.Run(() =>
				{
					while (!cancel)
					{
						const int millisecondsPerSecond = 1000;
						const int progressBarPeriodInMilliseconds = 500;
						DateTime now = DateTime.Now;
						int ms = (now.Second * millisecondsPerSecond) + now.Millisecond;
						int msfrac = ms % progressBarPeriodInMilliseconds;
						double pct = msfrac * (1.0 / (double)progressBarPeriodInMilliseconds);
						dialog.ProgressPct = pct;
						System.Threading.Thread.Sleep(200);
						if (dialog.WasCancelled)
						{
							cancel = true;
						}
					}
				});

				// create the actual task that does the work
				var whyMeTask = System.Threading.Tasks.Task.Run(() =>
				{
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.application_members.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.namespace_members.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mapifolder_members.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.items_members.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.itemsevents_event.itemadd.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.itemsevents_itemaddeventhandler.aspx
					//var inbox = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
					string searchFor = this.Application.Session.CurrentUser.Address;
					itemProcessed = true;
					whyMe = DetermineWhyMe(searchFor, mailItem, ref cancel, dialog);

					if (!cancel)
					{
						dialog.BeginInvoke(new System.Action(() => { dialog.Close(); }));
					}
				});

				NativeWindow mainWindow = new NativeWindow();
				mainWindow.AssignHandle(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);
				DialogResult dialogResult = dialog.ShowDialog(mainWindow);
				mainWindow.ReleaseHandle();
				cancel = true;
				whyMeTask.Wait();
				progressTask.Wait();

				if (dialog.WasCancelled)
				{
					return;
				}
			});

			WhyMeTask.Wait();

			hierarchy.AddRange(whyMe);
			return itemProcessed;
		}


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private bool DetermineWhyMeWithDialog1(Outlook.MailItem mailItem, ArrayList hierarchy)
		{
			bool cancel = false;
			bool itemProcessed = false;
			string[] whyMe = new string[0];

			// create the uber "WhyMe" task that will run in the NON-UI thread
			var WhyMeTask = System.Threading.Tasks.Task.Run(() =>
			{
				ProgressDialog dialog = new ProgressDialog();

				// create the progress bar task that merely mimics time
				var progressTask = System.Threading.Tasks.Task.Run(() =>
				{
					while (!cancel)
					{
						const int millisecondsPerSecond = 1000;
						const int progressBarPeriodInMilliseconds = 500;
						DateTime now = DateTime.Now;
						int ms = (now.Second * millisecondsPerSecond) + now.Millisecond;
						int msfrac = ms % progressBarPeriodInMilliseconds;
						double pct = msfrac * (1.0 / (double)progressBarPeriodInMilliseconds);
						dialog.ProgressPct = pct;
						System.Threading.Thread.Sleep(200);
						if (dialog.WasCancelled)
						{
							cancel = true;
						}
					}
				});

				// create the actual task that does the work
				var whyMeTask = System.Threading.Tasks.Task.Run(() =>
				{
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.application_members.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.namespace_members.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mapifolder_members.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.items_members.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.itemsevents_event.itemadd.aspx
					// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.itemsevents_itemaddeventhandler.aspx
					//var inbox = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox)
					string searchFor = this.Application.Session.CurrentUser.Address;
					itemProcessed = true;
					whyMe = DetermineWhyMe(searchFor, mailItem, ref cancel, dialog);

					if (!cancel)
					{
						dialog.BeginInvoke(new System.Action(() => { dialog.Close(); }));
					}
				});

				NativeWindow mainWindow = new NativeWindow();
				mainWindow.AssignHandle(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);
				DialogResult dialogResult = dialog.ShowDialog(mainWindow);
				mainWindow.ReleaseHandle();
				cancel = true;
				whyMeTask.Wait();
				progressTask.Wait();

				if (dialog.WasCancelled)
				{
					return;
				}
			});

			WhyMeTask.Wait();

			hierarchy.AddRange(whyMe);
			return itemProcessed;
		}


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private async Task<int> DoBusyWork(CancellationToken cancel, IProgress<double> progress)
		{
			int x = 124325267;
			int y = 987532465;
			int z = 1;
			var total = TimeSpan.FromSeconds(10);

			await Task.Run(() =>
			{
				var start = DateTime.Now;
				var end = start + total;

				var now = DateTime.Now;
				var inv_denominator = 100.0 / (end - start).TotalSeconds;

				while (now < end)
				{
					var complete = (now - start).TotalSeconds * inv_denominator;
					progress.Report(complete);
					cancel.ThrowIfCancellationRequested();

					z *= x;
					z *= y;

					Task.Delay(10).Wait();

					now = DateTime.Now;
				}
			});

			return z;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private async Task ShowDummyProgress()
		{
			//return Task.Run(() =>
			//{
			//});

			await Task.Delay(1);

			ProgressDialog dialog = new ProgressDialog();

			var cts = new CancellationTokenSource();
			var progress = new Progress<double>(percentage => { dialog.ProgressPct = percentage; });

			TaskCompletionSource<object> tcs = new System.Threading.Tasks.TaskCompletionSource<object>();

			dialog.FormClosed += delegate {
				cts.Cancel();
				tcs.SetResult(null);
			};

			dialog.Show();

			var r = await DoBusyWork(cts.Token, progress).ConfigureAwait(false);

			Logger.the.WriteLineFormat("Result: {0}", r);

			var letitgo = Task.Run(async () =>
			{
				await tcs.Task;
				nop();
			});


			return;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		Outlook.MAPIFolder FindSubFolderRecursive(Outlook.MAPIFolder root, string name, string alias = null)
		{
			Outlook.MAPIFolder found = null;

			if (alias != null)
			{
				foreach (Outlook.MAPIFolder subFolder in root.Folders)
				{
					Logger.the.WriteLineFormat("   Considering \"{0}\".", subFolder.FolderPath);

					foreach (var foldername in this.foldernames)
					{
						string thisfoldername = string.Format(foldername, alias);

						// debugging...
						if (subFolder.Name.IndexOf("Subway") != -1)
						{
							nop();
						}

						// debugging...
						if (subFolder.Name.IndexOf("Noise") != -1)
						{
							nop();
						}

						if (string.Compare(subFolder.Name, thisfoldername, true) == 0)
						{
							return subFolder;
						}

						if (string.Compare(subFolder.Name, alias, true) == 0)
						{
							return subFolder;
						}
					}

				}

				foreach (Outlook.MAPIFolder subFolder in root.Folders)
				{
					// debugging...
					if (subFolder.Name.IndexOf("Noise") != -1)
					{
						nop();
					}

					found = FindSubFolderRecursive(subFolder, name, alias);
					if (found != null)
					{
						return found;
					}
				}
			}

			foreach (Outlook.MAPIFolder subFolder in root.Folders)
			{
				Logger.the.WriteLineFormat("   Considering \"{0}\".", subFolder.FolderPath);

				if (string.Compare(subFolder.Name, name, true) == 0)
				{
					return subFolder;
				}
			}

			foreach (Outlook.MAPIFolder subFolder in root.Folders)
			{
				found = FindSubFolderRecursive(subFolder, name, alias);
				if (found != null)
				{
					break;
				}
			}

			return found;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private void CreateRule(dynamic mailItem)
		{
			if (this.RunRulesOnMailItem(mailItem))
			{
				return;
			}

			ArrayList whyMe = new ArrayList();
			bool itemProcessed = DetermineWhyMeWithDialog(mailItem, whyMe);

			if (!itemProcessed)
			{
				return;
			}

			if (whyMe.Count > 1)
			{
				string newRuleName = whyMe[whyMe.Count-2] as string;

				// find the "noise" folder
				var noise = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent.Folders["Noise"];
				nop(noise);

				// find the recipient
				// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_members.aspx
				//https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.recipients_members.aspx

				Outlook.Recipients recipients = mailItem.Recipients;
				Outlook.Recipient recipient = recipients.Cast<Outlook.Recipient>().Where(thisrecipient => thisrecipient.Name == newRuleName).FirstOrDefault();

				if (recipient == null)
				{
					MessageBox.Show("Recipient not found");
				}
				else
				{
					foreach (var pattern in patterns)
					{
						var match = Regex.Match(newRuleName, pattern);

						if (match.Success)
						{
							string pre = match.Groups["pre"].Value;
							string emailalias = match.Groups["alias"].Value;
							string post = match.Groups["post"].Value;

							if (emailalias.Length > 1)
							{
								newRuleName = emailalias.Substring(0,1).ToUpper() + emailalias.Substring(1).ToLower();
							}
							else
							{
								newRuleName = emailalias;
							}

							nop();

							break;
						}
						else
						{
							nop();
						}
					}

					Outlook.MAPIFolder currentFolder = this.Application.ActiveExplorer().CurrentFolder;
					currentFolder = noise;

					Outlook.MAPIFolder childFolder = null;
					try
					{
						childFolder = currentFolder.Folders[newRuleName];
					}
					catch { }

					if (childFolder == null)
					{
						childFolder = currentFolder.Folders.Add(newRuleName);
					}

					Rule newRule = new Rule(this._settings.Rules.Length + 1000, newRuleName, true,
						new Action[] { new MoveAction(childFolder.EntryID), },
						new RecipientCondition(recipient.Address, null));

					List<Rule> newRuleList = new List<Rule>(this._settings.Rules);
					newRuleList.Add(newRule);

					this._settings.Rules = newRuleList.ToArray();
					this.RunRulesOnMailItem(mailItem);

					Logger.the.WriteLineFormat("Rule created for \"{0}\".", newRuleName);
				}
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private void CreateSenderRule(dynamic item)
		{
			if (this.RunRulesOnMailItem(item))
			{
				return;
			}

			if (item is Outlook.MailItem)
			{
				Outlook.MailItem mailItem = item as Outlook.MailItem;

				string newRuleName = mailItem.Sender.Name;

				// find the "noise" folder
				var noise = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent.Folders["Noise"];
				nop(noise);

				Outlook.MAPIFolder currentFolder = this.Application.ActiveExplorer().CurrentFolder;
				currentFolder = noise;

				Outlook.MAPIFolder childFolder = null;
				try
				{
					childFolder = currentFolder.Folders[newRuleName];
				}
				catch { }

				if (childFolder == null)
				{
					childFolder = currentFolder.Folders.Add(newRuleName);
				}

				Rule newRule = new Rule(this._settings.Rules.Length + 1000, newRuleName, true,
					new Action[] { new MoveAction(childFolder.EntryID), },
					new SenderCondition(mailItem.Sender.Address, null));

				List<Rule> newRuleList = new List<Rule>(this._settings.Rules);
				newRuleList.Add(newRule);

				this._settings.Rules = newRuleList.ToArray();
				this.RunRulesOnMailItem(mailItem);

				Logger.the.WriteLineFormat("Rule created for \"{0}\".", newRuleName);

				return;
			}
			else if (item is Outlook.MeetingItem)
			{
				Outlook.MeetingItem mailItem = item as Outlook.MeetingItem;

				StringBuilder sb = new StringBuilder();
				sb.AppendFormat("Meeting: {0}\n", mailItem.SenderEmailAddress);
				MessageBox.Show(sb.ToString());

				return;
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private void CreateRuleWithFolder(dynamic mailItem)
		{
			if (this.RunRulesOnMailItem(mailItem))
			{
				return;
			}

			ArrayList whyMe = new ArrayList();
			bool itemProcessed = DetermineWhyMeWithDialog(mailItem, whyMe);

			if (!itemProcessed)
			{
				return;
			}

			if (whyMe.Count > 1)
			{
				string newRuleName = whyMe[whyMe.Count-2] as string;

				// find the recipient
				// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_members.aspx
				// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.recipients_members.aspx
				// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.recipient_members.aspx
				// Recipient Microsoft.Office.Interop.Outlook alias

				Outlook.Recipients recipients = mailItem.Recipients;
				Outlook.Recipient recipient = recipients.Cast<Outlook.Recipient>().Where(thisrecipient => thisrecipient.Name == newRuleName).FirstOrDefault();

				if (recipient == null)
				{
					MessageBox.Show("Recipient not found");
				}
				else
				{
					// try to find an existing folder
					//Outlook.MAPIFolder rootFolder = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
					Outlook.MAPIFolder rootFolder = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox).Parent;


					Outlook.MAPIFolder childFolder = FindSubFolderRecursive(rootFolder, newRuleName);

					if (childFolder == null)
					{
						const string PR_SMTP_ADDRESS = "http://schemas.microsoft.com/mapi/proptag/0x39FE001E";

						//childFolder = FindSubFolderRecursive(rootFolder, null, );
						Outlook.PropertyAccessor pa = recipient.PropertyAccessor;
						string smtpAddress = pa.GetProperty(PR_SMTP_ADDRESS).ToString();
						var atSymbolIndex = smtpAddress.IndexOf("@");
						if (atSymbolIndex != -1)
						{
							string emailAlias = smtpAddress.Substring(0, atSymbolIndex);
							childFolder = FindSubFolderRecursive(rootFolder, null, emailAlias);
						}

						nop();
					}

					if (childFolder == null)
					{
						string s;
						s = recipient.Name;
						s = recipient.Address;
						s = recipient.EntryID;
						s = recipient.AddressEntry.Address;
						s = recipient.AddressEntry.ID;
						s = recipient.AddressEntry.Name;
						// https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.contactitem_members.aspx
						var contact = recipient.AddressEntry.GetContact();

						if (contact != null)
						{
							s = contact.Account;
							s = contact.Account;
						}

						// PR_SUBJECT PR_SMTP_ADDRESS http://schemas.microsoft.com/mapi/ proptag
						// https://msdn.microsoft.com/en-us/library/office/ff184647.aspx

						string smtp = recipient.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E").ToString()	;
						string alias = smtp.Substring(0, smtp.IndexOf("@"));

						childFolder = FindSubFolderRecursive(rootFolder, alias);
						if (childFolder == null)
						{
							childFolder = Application.Session.PickFolder();
						}
					}

					if (childFolder != null)
					{
						Logger.the.WriteLineFormat("Putting item in \"{0}\".", childFolder.FolderPath);

						Rule newRule = new Rule(this._settings.Rules.Length + 1000, newRuleName, true,
							new Action[] { new MoveAction(childFolder.EntryID), },
							new RecipientCondition(recipient.Address, null));

						List<Rule> newRuleList = new List<Rule>(this._settings.Rules);
						newRuleList.Add(newRule);

						this._settings.Rules = newRuleList.ToArray();
						this.RunRulesOnMailItem(mailItem);

						Logger.the.WriteLineFormat("Rule created for \"{0}\".", newRuleName);
					}
				}
			}
			else
			{
				MessageBox.Show("Unknown. Maybe you were BCC'd.", "Unknown");
			}
		}


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private int FindFirstRuleIndexForItem(dynamic mailItem)
		{
			var meta = new Dictionary< string,string>();

			int index = 0;
			foreach (Rule rule in this._settings.Rules)
			{
				if (rule.Active)
				{
					if (rule.Test(mailItem, meta))
					{
						return index;
					}
				}

				++index;
			}

			return -1;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private int FindLastRuleIndexForItem(dynamic mailItem)
		{
			int index = 0;
			int lastFoundIndex = -1;
			int count = 0;

			var meta = new Dictionary<string,string>();

			foreach (Rule rule in this._settings.Rules)
			{
				if (rule.Active)
				{
					if (rule.Test(mailItem, meta))
					{
						++count;
						lastFoundIndex = index;
					}
				}

				++index;
			}

			MessageBox.Show("Num found: " + count.ToString(), "Custom Menu", MessageBoxButtons.OK);


			return lastFoundIndex;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private int _ruleIndexToBreakOn = 123;
		private bool _RunRulesOnMailItem(dynamic mailItem)
		{
			bool ruleWasRun = false;
			var meta = new Dictionary<string, string>();

			foreach (Rule rule in this._settings.Rules)
			{
				if (this._cancelRun)
				{
					break;
				}

				// debugging
				if (rule.Index == this._ruleIndexToBreakOn)
				{
					this.nop();
				}

				if (rule.Active)
				{
					if (rule.Test(mailItem, meta))
					{
						ruleWasRun = true;
						rule.InvokeActions(mailItem);

						if (rule.Final)
						{
							break;
						}
					}
				}
			}

			return ruleWasRun;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private bool RunRulesOnMailItem(Outlook.MailItem mailItem)
		{
			return this._RunRulesOnMailItem(mailItem);
		}

		//=============================================================================================================================================================================================p
		//=============================================================================================================================================================================================
		private bool RunRulesOnMailItem(Outlook.MeetingItem mailItem)
		{
			return this._RunRulesOnMailItem(mailItem);
		}


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		void CopyTextToClipboard(string text)
		{
			var dataObject = new System.Windows.Forms.DataObject();
			dataObject.SetData(System.Windows.Forms.DataFormats.Text, false, text);

			// Place the data object in the system clipboard.
			System.Windows.Forms.Clipboard.SetDataObject(dataObject, true);
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		IEnumerator<T> Cast<T>(IEnumerator iterator)
		{
			while (iterator.MoveNext())
			{
				yield return (T) iterator.Current;
			}
		}

		#endregion

		#region VSTO generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		#endregion

		// guarantee only one instance running
		private Mutex _exclusiveMutex;
		private bool _exclusiveMutexOwned = false;
		private const string _exclusiveMutexName = @"Global\RuleEx2025_exclusivitiy";

		private bool StartExclusive()
		{
			if (this._exclusiveMutex == null)
			{
				this._exclusiveMutex = new Mutex(false, _exclusiveMutexName);
				_exclusiveMutexOwned = this._exclusiveMutex.WaitOne(0);
				return _exclusiveMutexOwned;
			}

			return _exclusiveMutexOwned;
		}

		private void EndExclusive()
		{
			if (this._exclusiveMutex != null)
			{
				this._exclusiveMutex.ReleaseMutex();
				_exclusiveMutexOwned = false;
				this._exclusiveMutex = null;
			}
		}
	}
}


