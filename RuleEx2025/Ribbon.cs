using Microsoft.Office.Tools.Ribbon;

namespace RuleEx2025
{
	public partial class TaskpaneRibbon
	{
		private void TaskpaneRibbon_Load(object sender, RibbonUIEventArgs e)
		{
		}

		//=========================================================================================
		//=========================================================================================
		//=========================================================================================
		//=========================================================================================

		private void buttonRun_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnRun();
		}

		private void buttonSaveSettings_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnSaveSettings();
		}

		private void buttonLoadSettings_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnLoadSettings();
		}

		private void buttonRunAllInbox_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnRunAll();
		}

		private void buttonWhyMe_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnWhyMe();
		}

		private void buttonCreateRule_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnCreateRule();
		}

		private void buttonCreateSenderRule_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnCreateSenderRule();
		}

		private void buttonCreateRuleWithFolder_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnCreateRuleWithFolder();
		}

		private void buttonGoToItemSetting_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnThisSettings();
		}

		private void buttonGoToItemSetting2_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnThisSettings2();
		}

		private void buttonSettings_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnSettings();
		}

		private void buttonCancelRun_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnCancelRun();
		}

		private void buttonShowFolder_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnShowFolder();
		}

		private void buttonShowRule_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnShowRuleText();
		}

		private void buttonShowSender_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnShowSender();
		}

		private void buttonAddToBuildRule_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnAddToBuildRule();
		}

		private void buttonCheckInvalidFolders_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnCheckRulesForInvalidFolders();
		}

		private void buttonShowItemFolderPath_Click(object sender, RibbonControlEventArgs e)
		{
			//Globals.ThisAddIn.OnBtnShowItemFolderPath();
			Globals.ThisAddIn.OnBtnShowFolderPath();
		}

		//Globals.ThisAddIn.OnBtnCreateRuleWithFolderAsync();
		//Globals.ThisAddIn.OnBtnShowRecipients();
		//Globals.ThisAddIn.OnBtnWhyMeAsync();

		private void buttonShowRecipients_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnShowRecipients();
		}

		private void buttonPickUser_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnPickUser();
		}

		private void buttonPickFolder_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnPickFolder();
		}

		private void buttonSave_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnSave();
		}
	}
}
