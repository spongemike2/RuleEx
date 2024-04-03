namespace RuleEx2025
{
	partial class TaskpaneRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		public TaskpaneRibbon() : base(Globals.Factory.GetRibbonFactory())
		{
			InitializeComponent();
		}

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

		#region Component Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.tab1 = this.Factory.CreateRibbonTab();
			this.group1 = this.Factory.CreateRibbonGroup();
			this.group2 = this.Factory.CreateRibbonGroup();
			this.group3 = this.Factory.CreateRibbonGroup();
			this.group4 = this.Factory.CreateRibbonGroup();

			this.buttonRun					= this.Factory.CreateRibbonButton();
			this.buttonSaveSettings			= this.Factory.CreateRibbonButton();
			this.buttonLoadSettings			= this.Factory.CreateRibbonButton();
			this.buttonRunAllInbox			= this.Factory.CreateRibbonButton();
			this.buttonWhyMe				= this.Factory.CreateRibbonButton();
			this.buttonCreateRule			= this.Factory.CreateRibbonButton();
			this.buttonCreateSenderRule		= this.Factory.CreateRibbonButton();
			this.buttonCreateRuleWithFolder	= this.Factory.CreateRibbonButton();
			this.buttonGoToItemSetting		= this.Factory.CreateRibbonButton();
			this.buttonGoToItemSetting2		= this.Factory.CreateRibbonButton();
			this.buttonSettings				= this.Factory.CreateRibbonButton();
			this.buttonCancelRun			= this.Factory.CreateRibbonButton();
			this.buttonShowFolder			= this.Factory.CreateRibbonButton();
			this.buttonShowRule				= this.Factory.CreateRibbonButton();
			this.buttonShowSender			= this.Factory.CreateRibbonButton();
			this.buttonAddToBuildRule		= this.Factory.CreateRibbonButton();
			this.buttonCheckInvalidFolders	= this.Factory.CreateRibbonButton();
			this.buttonShowItemFolderPath	= this.Factory.CreateRibbonButton();
			this.buttonPickUser				= this.Factory.CreateRibbonButton();
			this.buttonPickFolder			= this.Factory.CreateRibbonButton();
			this.buttonSave					= this.Factory.CreateRibbonButton();
			this.buttonShowRecipients		= this.Factory.CreateRibbonButton();

			this.tab1.SuspendLayout();
			this.group1.SuspendLayout();
			this.group2.SuspendLayout();
			this.group3.SuspendLayout();
			this.group4.SuspendLayout();
			this.SuspendLayout();
			//
			// tab1
			//
			this.tab1.Groups.Add(this.group1);
			this.tab1.Groups.Add(this.group2);
			this.tab1.Groups.Add(this.group3);
			this.tab1.Groups.Add(this.group4);
			this.tab1.Label = "RuleEx 2025";
			this.tab1.Name = "tab1";

			//
			// group1
			//
			this.group1.Items.Add(this.buttonRun);
			this.group1.Items.Add(this.buttonRunAllInbox);
			this.group1.Items.Add(this.buttonSaveSettings);
			this.group1.Items.Add(this.buttonLoadSettings);

			this.group2.Items.Add(this.buttonShowItemFolderPath);
			this.group2.Items.Add(this.buttonWhyMe);
			this.group2.Items.Add(this.buttonCancelRun);
			this.group2.Items.Add(this.buttonShowFolder);
			this.group2.Items.Add(this.buttonShowRule);
			this.group2.Items.Add(this.buttonShowSender);
			this.group2.Items.Add(this.buttonShowRecipients);
			this.group2.Items.Add(this.buttonGoToItemSetting);
			this.group2.Items.Add(this.buttonGoToItemSetting2);
			this.group2.Items.Add(this.buttonSettings);

			this.group3.Items.Add(this.buttonCreateRule);
			this.group3.Items.Add(this.buttonCreateRuleWithFolder);
			this.group3.Items.Add(this.buttonCreateSenderRule);
			this.group3.Items.Add(this.buttonAddToBuildRule);
			this.group3.Items.Add(this.buttonSave);

			this.group4.Items.Add(this.buttonCheckInvalidFolders);
			this.group4.Items.Add(this.buttonPickUser);
			this.group4.Items.Add(this.buttonPickFolder);

			this.group1.Label = "Basic";
			this.group1.Name = "group1";

			this.group2.Label = "Search";
			this.group2.Name = "group2";

			this.group3.Label = "Utility";
			this.group3.Name = "group3";

			this.group4.Label = "Debugging";
			this.group4.Name = "group4";

			this.buttonRun.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonRunAllInbox.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonSaveSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonLoadSettings.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonCreateRule.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonCreateRuleWithFolder.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonCreateSenderRule.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
			this.buttonShowItemFolderPath.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;

			this.buttonRun.Label = "Run";
			this.buttonRun.Name = "Run";
			this.buttonRun.Description = "Run rules on the current selected items";
			this.buttonRun.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRun_Click);
			this.buttonRun.ShowImage = true;

			this.buttonSaveSettings.Label = "Save Settings";
			this.buttonSaveSettings.Name = "SaveSettings";
			this.buttonSaveSettings.Description = "Save the settings to disk";
			this.buttonSaveSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSaveSettings_Click);
			this.buttonSaveSettings.ShowImage = true;

			this.buttonLoadSettings.Label = "Load Settings";
			this.buttonLoadSettings.Name = "LoadSettings";
			this.buttonLoadSettings.Description = "Load the settings from disk";
			this.buttonLoadSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonLoadSettings_Click);
			this.buttonLoadSettings.ShowImage = true;

			this.buttonRunAllInbox.Label = "RunAll on Inbox";
			this.buttonRunAllInbox.Name = "RunAllInbox";
			this.buttonRunAllInbox.Description = "Run All Rules on All Inbox Items";
			this.buttonRunAllInbox.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonRunAllInbox_Click);
			this.buttonRunAllInbox.ShowImage = true;

			this.buttonWhyMe.Label = "Why Me";
			this.buttonWhyMe.Name = "WhyMe";
			this.buttonWhyMe.Description = "Determine why an email was sent to you";
			this.buttonWhyMe.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonWhyMe_Click);
			this.buttonWhyMe.ShowImage = true;

			this.buttonCreateRule.Label = "Create Rule";
			this.buttonCreateRule.Name = "CreateRule";
			this.buttonCreateRule.Description = "Create a \"Recipient Move-To\" rule for this email message";
			this.buttonCreateRule.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCreateRule_Click);
			this.buttonCreateRule.ShowImage = true;

			this.buttonCreateSenderRule.Label = "Create Sender Rule";
			this.buttonCreateSenderRule.Name = "CreateSenderRule";
			this.buttonCreateSenderRule.Description = "Create a \"Sender Move-To\" rule for this email message";
			this.buttonCreateSenderRule.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCreateSenderRule_Click);
			this.buttonCreateSenderRule.ShowImage = true;

			this.buttonCreateRuleWithFolder.Label = "Create Rule with Existing Folder";
			this.buttonCreateRuleWithFolder.Name = "CreateRuleWithFolder";
			this.buttonCreateRuleWithFolder.Description = "Create a \"Recipient Move-To\" rule for this email message";
			this.buttonCreateRuleWithFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCreateRuleWithFolder_Click);
			this.buttonCreateRuleWithFolder.ShowImage = true;

			this.buttonGoToItemSetting.Label = "Find Setting";
			this.buttonGoToItemSetting.Name = "GoToItemSetting";
			this.buttonGoToItemSetting.Description = "Go to the settings for this item";
			this.buttonGoToItemSetting.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGoToItemSetting_Click);
			this.buttonGoToItemSetting.ShowImage = true;

			this.buttonGoToItemSetting2.Label = "Find Setting (last)";
			this.buttonGoToItemSetting2.Name = "GoToItemSetting2";
			this.buttonGoToItemSetting2.Description = "Go to the settings for this item";
			this.buttonGoToItemSetting2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonGoToItemSetting2_Click);
			this.buttonGoToItemSetting2.ShowImage = true;

			this.buttonSettings.Label = "Settings";
			this.buttonSettings.Name = "Settings";
			this.buttonSettings.Description = "View or change settings";
			this.buttonSettings.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSettings_Click);
			this.buttonSettings.ShowImage = true;

			this.buttonCancelRun.Label = "Cancel Run";
			this.buttonCancelRun.Name = "CancelRun";
			this.buttonCancelRun.Description = "Cancels any running process";
			this.buttonCancelRun.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCancelRun_Click);
			this.buttonCancelRun.ShowImage = true;

			this.buttonShowFolder.Label = "Show Folder";
			this.buttonShowFolder.Name = "ShowFolder";
			this.buttonShowFolder.Description = "Show this folder's unique identifier";
			this.buttonShowFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonShowFolder_Click);
			this.buttonShowFolder.ShowImage = true;

			this.buttonShowRule.Label = "Show Rule Text";
			this.buttonShowRule.Name = "ShowRule";
			this.buttonShowRule.Description = "Show this mail item's rule text";
			this.buttonShowRule.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonShowRule_Click);
			this.buttonShowRule.ShowImage = true;

			this.buttonShowSender.Label = "Show Sender";
			this.buttonShowSender.Name = "ShowSender";
			this.buttonShowSender.Description = "Show this mail item's sender's unique identifier";
			this.buttonShowSender.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonShowSender_Click);
			this.buttonShowSender.ShowImage = true;

			this.buttonAddToBuildRule.Label = "Add To Build Rule";
			this.buttonAddToBuildRule.Name = "AddToBuildRule";
			this.buttonAddToBuildRule.Description = "Add the item to the build rules";
			this.buttonAddToBuildRule.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAddToBuildRule_Click);
			this.buttonAddToBuildRule.ShowImage = true;

			this.buttonCheckInvalidFolders.Label = "Check for Invalid Folders";
			this.buttonCheckInvalidFolders.Name = "CheckInvalidFolders";
			this.buttonCheckInvalidFolders.Description = "Check for Invalid Folders";
			this.buttonCheckInvalidFolders.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonCheckInvalidFolders_Click);
			this.buttonCheckInvalidFolders.ShowImage = true;

			this.buttonShowItemFolderPath.Label = "Show Folder Path";
			this.buttonShowItemFolderPath.Name = "ShowItemFolderPath";
			this.buttonShowItemFolderPath.Description = "Show the folder path the selected is in";
			this.buttonShowItemFolderPath.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonShowItemFolderPath_Click);
			this.buttonShowItemFolderPath.ShowImage = true;

			this.buttonPickUser.Label = "Pick Address Book Entry";
			this.buttonPickUser.Name = "PickUser";
			this.buttonPickUser.Description = "Pick an entry from the address book";
			this.buttonPickUser.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonPickUser_Click);
			this.buttonPickUser.ShowImage = true;

			this.buttonPickFolder.Label = "Pick Folder";
			this.buttonPickFolder.Name = "PickFolder";
			this.buttonPickFolder.Description = "Pick a folder using the built-in Outlook folder picker dialog";
			this.buttonPickFolder.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonPickFolder_Click);
			this.buttonPickFolder.ShowImage = true;

			this.buttonSave.Label = "Save HTML";
			this.buttonSave.Name = "Save";
			this.buttonSave.Description = "Save this item to HTML";
			this.buttonSave.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonSave_Click);
			this.buttonSave.ShowImage = true;

			this.buttonShowRecipients.Label = "Show Recipients";
			this.buttonShowRecipients.Name = "ShowRecipients";
			this.buttonShowRecipients.Description = "Show the recipients of this message";
			this.buttonShowRecipients.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonShowRecipients_Click);
			this.buttonShowRecipients.ShowImage = true;


			//
			// https://bert-toolkit.com/imagemso-list.html
			//

			this.buttonRun.OfficeImageId = "NewChessTool";
			this.buttonSaveSettings.OfficeImageId = "FileSave";
			this.buttonLoadSettings.OfficeImageId = "FileSave";
			this.buttonRunAllInbox.OfficeImageId = "SendReceiveAll";
			this.buttonWhyMe.OfficeImageId = "ShowContactPage";
			this.buttonCreateRule.OfficeImageId = "CreateMailRule";
			this.buttonCreateSenderRule.OfficeImageId = "CreateMailRule";
			this.buttonCreateRuleWithFolder.OfficeImageId = "CreateMailRule";
			this.buttonGoToItemSetting.OfficeImageId = "ZoomIn";
			this.buttonGoToItemSetting2.OfficeImageId = "ZoomOut";
			this.buttonSettings.OfficeImageId = "SetupClassicOffline"; // ToolboxGallery
			this.buttonCancelRun.OfficeImageId = "MergeViewClose";
			this.buttonShowFolder.OfficeImageId = "NewCategoryFolder";
			this.buttonShowRule.OfficeImageId = "MessageProperties";
			this.buttonShowSender.OfficeImageId = "ContactPictureMenu";//"HappyFace";//"MeetingsToolNext";
			this.buttonAddToBuildRule.OfficeImageId = "ContentControlBuildingBlockGallery";
			this.buttonCheckInvalidFolders.OfficeImageId = "TableTestValidationRules";
			this.buttonShowItemFolderPath.OfficeImageId = "HeaderFooterFilePathInsert";
			this.buttonPickUser.OfficeImageId = "ShowMembersPage";
			this.buttonPickFolder.OfficeImageId = "Folder";
			this.buttonSave.OfficeImageId = "RecordsSaveRecord";
			this.buttonShowRecipients.OfficeImageId = "ShowMembersPage";



			//
			// TaskpaneRibbon
			//
			this.Name = "TaskpaneRibbon";
			this.RibbonType = "Microsoft.Outlook.Explorer";
			this.Tabs.Add(this.tab1);
			this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.TaskpaneRibbon_Load);

			this.tab1.ResumeLayout(false);
			this.group1.ResumeLayout(false);
			this.group2.ResumeLayout(false);
			this.group3.ResumeLayout(false);
			this.group4.ResumeLayout(false);

			this.tab1.PerformLayout();
			this.group1.PerformLayout();
			this.group2.PerformLayout();
			this.group3.PerformLayout();
			this.group4.PerformLayout();

			this.ResumeLayout(false);
		}

		#endregion

		internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
		internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;

		// https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.tools.ribbon?view=vsto-2017

		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRun;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSaveSettings;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonLoadSettings;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonRunAllInbox;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonWhyMe;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCreateRule;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCreateSenderRule;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCreateRuleWithFolder;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGoToItemSetting;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonGoToItemSetting2;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSettings;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCancelRun;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonShowFolder;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonShowRule;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonShowSender;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonAddToBuildRule;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonCheckInvalidFolders;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonShowItemFolderPath;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonPickUser;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonPickFolder;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonSave;
		internal Microsoft.Office.Tools.Ribbon.RibbonButton buttonShowRecipients;
	}

	partial class ThisRibbonCollection
	{
		internal TaskpaneRibbon TaskpaneRibbon
		{
			get { return this.GetRibbon<TaskpaneRibbon>(); }
		}
	}
}







