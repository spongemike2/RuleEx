using Microsoft.Office.Tools.Ribbon;

namespace RuleEx2025
{
	public partial class MailReadRibbon
	{
		private void MailReadRibbon_Load(object sender, RibbonUIEventArgs e)
		{
		}

		private void button1_Click(object sender, RibbonControlEventArgs e)
		{
			Globals.ThisAddIn.OnBtnShowItemFolderPath();
		}
	}
}
