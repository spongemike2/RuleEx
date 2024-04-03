using System;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RuleEx2025
{
	public partial class GroupDialog : Form
	{
		private Outlook.Application _application;
		private ConditionGroup _group;

		public GroupDialog()
		{
			InitializeComponent();
		}

		public GroupDialog(Outlook.Application application, ConditionGroup cg) : this()
		{
			this._application = application;
			this._group = cg;
		}

		private void groupBox1_Enter(object sender, EventArgs e)
		{
		}

		private void GroupDialog_Load(object sender, EventArgs e)
		{
			this.SuspendLayout();
			this.radioButtonAnd.Checked = this._group.Operator == ConditionGroup.GroupingOperator.And;
			this.radioButtonOr.Checked = this._group.Operator == ConditionGroup.GroupingOperator.Or;
			this.ResumeLayout(true);
		}

		protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
		{
			if (keyData == Keys.Escape)
			{
				this.Close();
				return true;
			}

			return base.ProcessCmdKey(ref msg, keyData);
		}
	}
}
