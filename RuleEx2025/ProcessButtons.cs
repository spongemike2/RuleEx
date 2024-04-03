using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
		//=============================================================================================================================================================================================
		//
		//
		//      ######## ##     ## ########  ##        #######  ########  ######## ########       ########  ##     ## ######## ########  #######  ##    ##  ######
		//      ##        ##   ##  ##     ## ##       ##     ## ##     ## ##       ##     ##      ##     ## ##     ##    ##       ##    ##     ## ###   ## ##    ##
		//      ##         ## ##   ##     ## ##       ##     ## ##     ## ##       ##     ##      ##     ## ##     ##    ##       ##    ##     ## ####  ## ##
		//      ######      ###    ########  ##       ##     ## ########  ######   ########       ########  ##     ##    ##       ##    ##     ## ## ## ##  ######
		//      ##         ## ##   ##        ##       ##     ## ##   ##   ##       ##   ##        ##     ## ##     ##    ##       ##    ##     ## ##  ####       ##
		//      ##        ##   ##  ##        ##       ##     ## ##    ##  ##       ##    ##       ##     ## ##     ##    ##       ##    ##     ## ##   ### ##    ##
		//      ######## ##     ## ##        ########  #######  ##     ## ######## ##     ##      ########   #######     ##       ##     #######  ##    ##  ######
		//
		//
		// Explorer Buttons
		//=============================================================================================================================================================================================

		public void OnBtnLoadSettings()
		{
			this._settings = Settings.Load(this._settingsFile);
		}

		public void OnBtnSaveSettings()
		{
			this._settings.Save(this._settingsFile);
		}

		public void OnBtnPickUser()
		{
			Outlook.SelectNamesDialog snd = Application.Session.GetSelectNamesDialog();
			snd.SetDefaultDisplayMode(Outlook.OlDefaultSelectNamesDisplayMode.olDefaultSingleName);
			var result = snd.Display();

			if (result)
			{
				if (snd.Recipients.Count == 1)
				{
					// why is the first one index "1" and not index "0"?!?!
					Outlook.Recipient recipient = snd.Recipients[1];

					Outlook.AddressEntry addressEntry = recipient.AddressEntry;
					string addressEntryAddress = addressEntry.Address;
					string address = recipient.Address;

					this.nop(recipient);
				}
			}
		}

		public void OnBtnPickFolder()
		{
			var folder = Application.Session.PickFolder();
			this.nop(folder);
		}

		public void OnBtnCancelRun()
		{
			this._cancelRun = true;
		}

		public void OnBtnSave()
		{
			Outlook.Selection selection = this.Application.ActiveExplorer().Selection;
			object o = selection.Cast<object>().FirstOrDefault();

			Outlook.MailItem mailItem = o as Outlook.MailItem;
			if (mailItem != null)
			{
				SaveFileDialog saveFileDialog = new SaveFileDialog();
				saveFileDialog.Filter = "(Html Files)|*.htm|(Html Files)|*.html";
				saveFileDialog.Title = "Save Email File";
				DialogResult result = saveFileDialog.ShowDialog();

				if (result == DialogResult.OK)
				{
					this.nop();
					System.IO.File.WriteAllText(saveFileDialog.FileName, mailItem.HTMLBody);
				}
			}
		}


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnSettings()
		{
			SettingsDialog dialog = new SettingsDialog(this._settings, this.Application);
			NativeWindow mainWindow = new NativeWindow();
			mainWindow.AssignHandle(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);
			DialogResult dialogResult = dialog.ShowDialog(mainWindow);
			if (dialogResult == DialogResult.OK)
			{
				this._settings = dialog.Settings;
			}
			mainWindow.ReleaseHandle();
		}


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnThisSettings()
		{
			Outlook.Selection selection = this.Application.ActiveExplorer().Selection;
			object o = selection.Cast<object>().FirstOrDefault();

			if (o != null)
			{
				int index = FindFirstRuleIndexForItem(o);

				SettingsDialog dialog = new SettingsDialog(this._settings, this.Application, index);
				NativeWindow mainWindow = new NativeWindow();
				mainWindow.AssignHandle(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);
				DialogResult dialogResult = dialog.ShowDialog(mainWindow);
				if (dialogResult == DialogResult.OK)
				{
					this._settings = dialog.Settings;
				}
				mainWindow.ReleaseHandle();
			}
			//dialogResult;
		}


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnThisSettings2()
		{
			Outlook.Selection selection = this.Application.ActiveExplorer().Selection;
			object o = selection.Cast<object>().FirstOrDefault();

			if (o != null)
			{
				int index = FindLastRuleIndexForItem(o);

				SettingsDialog dialog = new SettingsDialog(this._settings, this.Application, index);
				NativeWindow mainWindow = new NativeWindow();
				mainWindow.AssignHandle(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);
				DialogResult dialogResult = dialog.ShowDialog(mainWindow);
				if (dialogResult == DialogResult.OK)
				{
					this._settings = dialog.Settings;
				}
				mainWindow.ReleaseHandle();
			}
			//dialogResult;
		}


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnWhyMe()
		{
			Outlook.Selection selection = this.Application.ActiveExplorer().Selection;
			object o = selection.Cast<object>().FirstOrDefault();

			// see what other kind of object "o" could be
			if (o is Outlook.AppointmentItem)
			{
				nop();
			}
			if (o is Outlook.ContactItem)
			{
				nop();
			}
			if (o is Outlook.DistListItem)
			{
				nop();
			}
			if (o is Outlook.DocumentItem)
			{
				nop();
			}
			if (o is Outlook.JournalItem)
			{
				nop();
			}
			if (o is Outlook.MailItem)
			{
				nop();
			}
			if (o is Outlook.MeetingItem)
			{
				nop();
			}
			if (o is Outlook.MobileItem)
			{
				nop();
			}
			if (o is Outlook.NoteItem)
			{
				nop();
			}
			if (o is Outlook.PostItem)
			{
				nop();
			}
			if (o is Outlook.RemoteItem)
			{
				nop();
			}
			if (o is Outlook.ReportItem)
			{
				nop();
			}
			if (o is Outlook.SharingItem)
			{
				nop();
			}
			if (o is Outlook.StorageItem)
			{
				nop();
			}
			if (o is Outlook.TaskItem)
			{
				nop();
			}
			if (o is Outlook.TaskRequestAcceptItem)
			{
				nop();
			}
			if (o is Outlook.TaskRequestDeclineItem)
			{
				nop();
			}
			if (o is Outlook.TaskRequestItem)
			{
				nop();
			}
			if (o is Outlook.TaskRequestUpdateItem)
			{
				nop();
			}


			var mailItem = o as Outlook.MailItem;
			if (mailItem != null)
			{
				ArrayList whyMe = new ArrayList();
				bool itemProcessed = DetermineWhyMeWithDialog(mailItem, whyMe);

				if (!itemProcessed)
				{
					return;
				}

				if (whyMe.Count == 0)
				{
					if (itemProcessed)
					{
						MessageBox.Show("Unknown. Maybe you were BCC'd.", "Unknown");
					}
					else
					{
						MessageBox.Show("No item in selection to check.", "Nothing to check");
					}
				}
				else if (whyMe.Count == 2)
				{
					MessageBox.Show("It was sent DIRECTLY to you, silly", "Found you");
				}
				else
				{
					StringBuilder sb = new StringBuilder();

					for (int i = 1; i < whyMe.Count; i++)
					{
						sb.Append(' ', 4 * i);
						sb.AppendFormat("{0}\n", whyMe[whyMe.Count - i - 1]);
					}

					MessageBox.Show(sb.ToString(), "Found you");
				}
			}
			else
			{
				var meetingItem = o as Outlook.MeetingItem;
				if (meetingItem != null)
				{
					ArrayList whyMe = new ArrayList();
					bool itemProcessed = DetermineWhyMeWithDialog(meetingItem, whyMe);

					if (!itemProcessed)
					{
						return;
					}

					if (whyMe.Count == 0)
					{
						if (itemProcessed)
						{
							MessageBox.Show("Unknown. Maybe you were BCC'd.", "Unknown");
						}
						else
						{
							MessageBox.Show("No item in selection to check.", "Nothing to check");
						}
					}
					else if (whyMe.Count == 2)
					{
						MessageBox.Show("It was sent DIRECTLY to you, silly", "Found you");
					}
					else
					{
						StringBuilder sb = new StringBuilder();

						for (int i = 1; i < whyMe.Count; i++)
						{
							sb.Append(' ', 4 * i);
							sb.AppendFormat("{0}\n", whyMe[whyMe.Count - i - 1]);
						}

						MessageBox.Show(sb.ToString(), "Found you");
					}
				}
				else
				{
				}
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnWhyMeAsync()
		{
			var foo = this.ShowDummyProgress();

			//foo.Wait();

			nop();
			//task.Wait();
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnCreateRule()
		{
			Outlook.Selection selection = this.Application.ActiveExplorer().Selection;

			foreach (object item in selection)
			{
				Outlook.MailItem mailItem = item as Outlook.MailItem;

				if (mailItem == null)
				{
					Outlook.MeetingItem meetingItem = item as Outlook.MeetingItem;
					if (meetingItem == null)
					{
						continue;
					}
					else
					{
						this.CreateRule(meetingItem);
					}
				}
				else
				{
					this.CreateRule(mailItem);
				}
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnCreateSenderRule()
		{
			Outlook.Selection selection = this.Application.ActiveExplorer().Selection;

			foreach (object item in selection)
			{
				Outlook.MailItem mailItem = item as Outlook.MailItem;

				if (mailItem == null)
				{
					Outlook.MeetingItem meetingItem = item as Outlook.MeetingItem;
					if (meetingItem == null)
					{
						continue;
					}
					else
					{
						this.CreateSenderRule(meetingItem);
					}
				}
				else
				{
					this.CreateSenderRule(mailItem);
				}
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnCreateRuleWithFolder()
		{
			Outlook.Selection selection = this.Application.ActiveExplorer().Selection;

			foreach (object item in selection)
			{
				Outlook.MailItem mailItem = item as Outlook.MailItem;

				if (mailItem == null)
				{
					Outlook.MeetingItem meetingItem = item as Outlook.MeetingItem;
					if (meetingItem == null)
					{
						continue;
					}
					else
					{
						this.CreateRuleWithFolder(meetingItem);
					}
				}
				else
				{
					this.CreateRuleWithFolder(mailItem);
				}
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnShowFolder()
		{
			Outlook.MAPIFolder folder = this.Application.ActiveExplorer().CurrentFolder;

			StringBuilder sb = new StringBuilder();
			sb.AppendFormat("<MoveAction FolderName=\"{0}\" FolderPath=\"{1}\" />", folder.EntryID, folder.FolderPath);
			Logger.the.WriteLineFormat("<MoveAction FolderName=\"{0}\" FolderPath=\"{1}\" />", folder.EntryID, folder.FolderPath);
			this.CopyTextToClipboard(sb.ToString());
			MessageBox.Show(sb.ToString());
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnShowRuleText()
		{
			Outlook.Selection selection = this.Application.ActiveExplorer().Selection;
			StringBuilder sb = new StringBuilder();

			foreach (var item in selection)
			{
				if (item is Outlook.MailItem)
				{
					Outlook.MailItem mailItem = item as Outlook.MailItem;

					foreach (Outlook.Recipient recipient in mailItem.Recipients)
					{
						sb.AppendFormat("<Rule Not=\"false\" Operator=\"Or\" Name=\"\" Active=\"true\" Final=\"true\" >\n");
						sb.AppendFormat("  <RecipientCondition Not=\"false\" Recipient=\"{0}\" />\n", recipient.Address);
						sb.AppendFormat("</Rule>\n");

						Logger.the.WriteLineFormat("<RecipientCondition Not=\"false\" Recipient=\"{0}\" />", recipient.Address);
						Logger.the.WriteLineFormat("<Recipient>{0}</Recipient>", recipient.Address);
					}
				}
				else if (item is Outlook.MeetingItem)
				{
					Outlook.MeetingItem mailItem = item as Outlook.MeetingItem;

					foreach (Outlook.Recipient recipient in mailItem.Recipients)
					{
						sb.AppendFormat("<Rule Not=\"false\" Operator=\"Or\" Name=\"\" Active=\"true\" Final=\"true\" >\n");
						sb.AppendFormat("  <RecipientCondition Not=\"false\" Recipient=\"{0}\" />\n", recipient.Address);
						sb.AppendFormat("</Rule>\n");
						Logger.the.WriteLineFormat("<RecipientCondition Not=\"false\" Recipient=\"{0}\" />", recipient.Address);
						Logger.the.WriteLineFormat("<Recipient>{0}</Recipient>", recipient.Address);
					}
				}
			}

			this.CopyTextToClipboard(sb.ToString());
			MessageBox.Show(sb.ToString());
		}


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnShowRecipients()
		{
			Outlook.Selection selection = this.Application.ActiveExplorer().Selection;
			StringBuilder sb = new StringBuilder();

			foreach (var item in selection)
			{
				if (item is Outlook.MailItem)
				{
					Outlook.MailItem mailItem = item as Outlook.MailItem;

					foreach (Outlook.Recipient recipient in mailItem.Recipients)
					{
						sb.AppendFormat("<RecipientCondition Not=\"false\" Recipient=\"{0}\" />\n", recipient.Address);
						Logger.the.WriteLineFormat("<RecipientCondition Not=\"false\" Recipient=\"{0}\" />", recipient.Address);
						Logger.the.WriteLineFormat("<Recipient>{0}</Recipient>", recipient.Address);
					}
				}
				else if (item is Outlook.MeetingItem)
				{
					Outlook.MeetingItem mailItem = item as Outlook.MeetingItem;

					foreach (Outlook.Recipient recipient in mailItem.Recipients)
					{
						sb.AppendFormat("<RecipientCondition Not=\"false\" Recipient=\"{0}\" />\n", recipient.Address);
						Logger.the.WriteLineFormat("<RecipientCondition Not=\"false\" Recipient=\"{0}\" />", recipient.Address);
						Logger.the.WriteLineFormat("<Recipient>{0}</Recipient>", recipient.Address);
					}
				}
			}

			this.CopyTextToClipboard(sb.ToString());
			MessageBox.Show(sb.ToString());
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnCheckRulesForInvalidFolders()
		{
			foreach (Rule rule in this._settings.Rules)
			{
				foreach (var action in rule.Actions)
				{
					if (action is MoveAction)
					{
						MoveAction maction = action as MoveAction;

						bool valid = false;
						try
						{
							maction.Folder = theApplication.Session.GetFolderFromID(maction.FolderName);
							string folderPath = maction.Folder.FolderPath;
							maction.FolderPath = folderPath;
							Logger.the.WriteLineFormat("The rule \"{0}\" has a valid folder.", rule.Name);
							valid = true;
						}
						catch
						{
						}

						if (valid)
						{
							continue;
						}

						Logger.the.WriteLineFormat("The rule \"{0}\" has an INVALID folder.", rule.Name);

						if (string.IsNullOrEmpty(maction.FolderPath))
						{
							this.nop();
						}
						else
						{
							// replace the folder with the matching one...
							string[] folderParts = maction.FolderPath.Split('\\');

							Outlook.MAPIFolder thisFolder = this.Application.Session.DefaultStore.GetRootFolder();

							bool rootFolder = true;

							foreach (string part in folderParts)
							{
								if (string.IsNullOrEmpty(part))
								{
									continue;
								}

								if (rootFolder)
								{
									if (thisFolder.Name == part)
									{
										rootFolder = false;
										continue;
									}
								}
								else
								{
									try
									{
										Outlook.MAPIFolder child = thisFolder.Folders[part];
										thisFolder = child;
									}
									catch
									{
										thisFolder = thisFolder.Folders.Add(part);
										MessageBox.Show(string.Format("Creating a new folder: {0}", thisFolder.FolderPath));
									}
								}
							}

							maction.Folder = thisFolder;
							maction.FolderName = thisFolder.EntryID;
							maction.FolderPath = thisFolder.FolderPath;
						}
					}
				}
			}
		}


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnAddToBuildRule()
		{
			// find the build rule
			Outlook.Selection selection = this.Application.ActiveExplorer().Selection;

			BuildRule br = null;

			foreach (Rule rule in this._settings.Rules)
			{
				// debugging
				br = rule as BuildRule;
				if (br != null)
				{
					break;
				}

			}

			if (br != null)
			{
				foreach (var item in selection)
				{
					Outlook.MailItem mailItem = item as Outlook.MailItem;

					if (mailItem == null)
					{
						continue;
					}

					if (this.RunRulesOnMailItem(mailItem))
					{
						continue;
					}

					// get why we have it
					ArrayList whyMe = new ArrayList();
					bool itemProcessed = DetermineWhyMeWithDialog(mailItem, whyMe);

					if (!itemProcessed)
					{
						continue;
					}

					if (whyMe.Count > 1)
					{
						ArrayList al;

						if (br.DistributionLists == null)
						{
							al = new ArrayList();
						}
						else
						{
							al = new ArrayList(br.DistributionLists);
						}
						al.Add(whyMe[whyMe.Count-1]);
						br.DistributionLists = al.ToArray(typeof(string)) as string[];
					}
				}
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnShowSender()
		{
			Outlook.Selection selection = this.Application.ActiveExplorer().Selection;
			StringBuilder sb = new StringBuilder();

			foreach (var item in selection)
			{
				if (item is Outlook.MailItem)
				{
					Outlook.MailItem mailItem = item as Outlook.MailItem;
					sb.AppendFormat("<SenderCondition Not=\"false\" Sender=\"{0}\" />\n", mailItem.Sender.Address);
				}
				else if (item is Outlook.MeetingItem)
				{
					Outlook.MeetingItem mailItem = item as Outlook.MeetingItem;
					sb.AppendFormat("<SenderCondition Not=\"false\" Sender=\"{0}\" />\n", mailItem.SenderEmailAddress);
				}
			}

			this.CopyTextToClipboard(sb.ToString());
			MessageBox.Show(sb.ToString());
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnRun()
		{
			System.GC.Collect();

			// we have the mutex
			Outlook.Selection selection = this.Application.ActiveExplorer().Selection;

			// create the uber "WhyMe" task that will run in the NON-UI thread
			var WhyMeTask = System.Threading.Tasks.Task.Run(() =>
			{
				if (this._runMutex.WaitOne(1))
				{
					this._cancelRun = false;

					bool showDialog = selection.Count > 1;

					//
					// convert to an IEnumerable<>
					//
					var sel = new List<object>(selection.Count);
					foreach (var item in selection)
					{
						if (this._cancelRun)
						{
							break;
						}

						sel.Add(item);
					}

					Parallel.ForEach(sel, item =>
					{
						if (!this._cancelRun)
						{
							if (item is Outlook.MailItem)
							{
								Outlook.MailItem mailItem = item as Outlook.MailItem;
								this.RunRulesOnMailItem(mailItem);
							}
							else if (item is Outlook.MeetingItem)
							{
								Outlook.MeetingItem mailItem = item as Outlook.MeetingItem;
								this.RunRulesOnMailItem(mailItem);
							}
						}
					});

					this._runMutex.ReleaseMutex();


					if (showDialog)
					{
						MessageBox.Show("Done running");
					}
				}
				else
				{
					// we don't have the mutex
					MessageBox.Show("Another run is already going");
				}
			});
		}


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void OnBtnRunAll()
		{
			System.GC.Collect();

			// create the uber "WhyMe" task that will run in the NON-UI thread
			var WhyMeTask = System.Threading.Tasks.Task.Run(() =>
			{
				if (this._runMutex.WaitOne(1))
				{
					this._cancelRun = false;

					var inbox = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
					var inboxItems = inbox.Items;

					// since processing the enumerator for the inbox items while moving them around seems to affect the enumerator,
					// we make a copy of all the items before we process it

					//
					// convert to an IEnumerable<>
					//
					var sel = new List<object>(inboxItems.Count);
					foreach (var item in inboxItems)
					{
						if (this._cancelRun)
						{
							break;
						}

						sel.Add(item);
					}

					Parallel.ForEach(sel, item =>
					{
						if (!this._cancelRun)
						{
							if (item is Outlook.MailItem)
							{
								this.RunRulesOnMailItem(item as Outlook.MailItem);
							}
							else if (item is Outlook.MeetingItem)
							{
								Outlook.MeetingItem mailItem = item as Outlook.MeetingItem;
								this.RunRulesOnMailItem(mailItem);
							}
						}
					});

					this._runMutex.ReleaseMutex();
					MessageBox.Show("Done running");
				}
				else
				{
					// we don't have the mutex
					MessageBox.Show("Another run is already going");
				}
			});
		}

		public void OnBtnShowFolderPath()
		{
			Outlook.Selection selection = this.Application.ActiveExplorer().Selection;

			foreach (var item in selection)
			{
				Outlook.MailItem mailItem = item as Outlook.MailItem;
				if (mailItem != null)
				{
					Outlook.MAPIFolder folder = mailItem.Parent;
					//this.CopyTextToClipboard(folder.FolderPath);
					MessageBox.Show(folder.FolderPath);

					return;
				}
			}
		}


		//=============================================================================================================================================================================================
		//
		//
		//      ##     ##    ###    #### ##            #### ######## ######## ##     ##      ########  ##     ## ######## ########  #######  ##    ##  ######
		//      ###   ###   ## ##    ##  ##             ##     ##    ##       ###   ###      ##     ## ##     ##    ##       ##    ##     ## ###   ## ##    ##
		//      #### ####  ##   ##   ##  ##             ##     ##    ##       #### ####      ##     ## ##     ##    ##       ##    ##     ## ####  ## ##
		//      ## ### ## ##     ##  ##  ##             ##     ##    ######   ## ### ##      ########  ##     ##    ##       ##    ##     ## ## ## ##  ######
		//      ##     ## #########  ##  ##             ##     ##    ##       ##     ##      ##     ## ##     ##    ##       ##    ##     ## ##  ####       ##
		//      ##     ## ##     ##  ##  ##             ##     ##    ##       ##     ##      ##     ## ##     ##    ##       ##    ##     ## ##   ### ##    ##
		//      ##     ## ##     ## #### ########      ####    ##    ######## ##     ##      ########   #######     ##       ##     #######  ##    ##  ######
		//
		//
		// Mail Item Buttons
		//=============================================================================================================================================================================================
		public void OnBtnShowItemFolderPath()
		{
			//
			// get the active inspector, i.e., the mail message we're looking at
			//
			// https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.inspector?view=outlook-pia
			var activeInspector = this.Application.ActiveInspector();
			if (activeInspector != null)
			{
				if (activeInspector.CurrentItem != null)
				{
					//
					// try mail item
					//
					Outlook.MailItem mailItem = activeInspector.CurrentItem as Outlook.MailItem;

					if (mailItem != null)
					{
						Outlook.MAPIFolder folder = mailItem.Parent;
						MessageBox.Show(folder.FolderPath);
					}
					else
					{
						//
						// try meeting item
						//
						var meetingItem = activeInspector.CurrentItem as Outlook.MeetingItem;
						Outlook.MAPIFolder folder = meetingItem.Parent;
						MessageBox.Show(folder.FolderPath);
					}
				}
			}
		}
	}
}
