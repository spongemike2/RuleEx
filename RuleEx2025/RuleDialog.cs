using System;
using System.Windows.Forms;

namespace RuleEx2025
{
	public partial class RuleDialog : Form
	{

		public RuleDialog()
		{
			InitializeComponent();
		}

		private Rule _rule;
		public Rule Rule
		{
			get
			{
				return this._rule;
			}
		}

		public RuleDialog(Rule rule) : this()
		{
			this._rule = rule.Clone();
		}

		private void RuleDialog_Load(object sender, EventArgs e)
		{
			this.nameTextBox.Text = this._rule.Name;
			this.lastRunTextBox.Text = this._rule.RunCount == 0 ? "Never" : this._rule.LastRun.ToString();
			this.runCountTextBox.Text = this._rule.RunCount.ToString();
			this.indexTextBox.Text = this._rule.Index.ToString();

			this.activeCheckBox.Checked = this._rule.Active;
			this.finalCheckBox.Checked = this._rule.Final;
			this.orOperatorButton.Checked = this._rule.Operator == ConditionGroup.GroupingOperator.Or;
			this.andOperatorButton.Checked = this._rule.Operator == ConditionGroup.GroupingOperator.And;
		}

		private void okButton_Click(object sender, EventArgs e)
		{
			this._rule.Name = this.nameTextBox.Text;
			this._rule.Active = this.activeCheckBox.Checked;
			this._rule.Final = this.finalCheckBox.Checked;
			this._rule.Operator = this.orOperatorButton.Checked ? ConditionGroup.GroupingOperator.Or : ConditionGroup.GroupingOperator.And;
		}

		private void operatorGroup_Enter(object sender, EventArgs e)
		{

		}
	}
}
