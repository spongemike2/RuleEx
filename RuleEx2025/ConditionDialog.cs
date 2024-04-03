using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RuleEx2025
{
	public partial class ConditionDialog : Form
	{
		private Outlook.Application _application;
		Microsoft.Office.Interop.Outlook.Recipient _recipient;
		Font _boldFont;
		Font _regFont;
		Control lastRadioButton = null;

		private Condition _condition;

		public RecipientCondition RecipientCondition
		{
			get
			{
				return this._condition as RecipientCondition;
			}
		}

		public SenderCondition SenderCondition
		{
			get
			{
				return this._condition as SenderCondition;
			}
		}

		public SubjectCondition SubjectCondition
		{
			get
			{
				return this._condition as SubjectCondition;
			}
		}

		public BodyCondition BodyCondition
		{
			get
			{
				return this._condition as BodyCondition;
			}
		}

		private void nop(Object o=null)
		{
		}

		public ConditionDialog()
		{
			InitializeComponent();
		}

		public ConditionDialog(Outlook.Application application, RecipientCondition rc) : this()
		{
			this._application = application;

			this.SuspendLayout();
			this.Text = "Recipient Condition Settings";
			this._condition = Clone(rc) as RecipientCondition;
			this.Init();
			this.ResumeLayout(true);
		}


		public ConditionDialog(Outlook.Application application, SenderCondition sc) : this()
		{
			this._application = application;

			this.SuspendLayout();
			this.Text = "Sender Condition Settings";
			this._condition = Clone(sc) as SenderCondition;
			this.Init();
			this.ResumeLayout(true);
		}

		public ConditionDialog(Outlook.Application application, SubjectCondition sc) : this()
		{
			this._application = application;

			this.SuspendLayout();
			this.Text = "Subject Condition Settings";
			this._condition = Clone(sc) as SubjectCondition;
			this.buttonPickUser.Visible = false;
			this.groupBox1.Visible = false;
			this.recipientNameLabel.Text = "Regex";
			this.nameTextBox.Text = sc.Regex;

			this.Init();
			this.ResumeLayout(true);
		}


		public ConditionDialog(Outlook.Application application, BodyCondition bc) : this()
		{
			this._application = application;

			this.SuspendLayout();
			this.Text = "Body Condition Settings";
			this._condition = Clone(bc) as BodyCondition;
			this.buttonPickUser.Visible = false;
			this.groupBox1.Visible = false;
			this.recipientNameLabel.Text = "Regex";
			this.nameTextBox.Text = bc.Regex;

			this.Init();
			this.ResumeLayout(true);
		}



		private Microsoft.Office.Interop.Outlook.Recipient GetRecipient()
		{
			Microsoft.Office.Interop.Outlook.Recipient recipient = null;

			if (this._condition is SenderCondition)
			{
				SenderCondition c = this._condition as SenderCondition;
				recipient = this._application.Session.CreateRecipient(c.Sender);
			}
			else if (this._condition is RecipientCondition)
			{
				RecipientCondition c = this._condition as RecipientCondition;
				recipient = this._application.Session.CreateRecipient(c.Recipient);
			}

			return recipient;
		}

		private string GetRegex()
		{
			if (this._condition is SenderCondition)
			{
				SenderCondition c = this._condition as SenderCondition;
				return c.Regex;
			}
			else if (this._condition is RecipientCondition)
			{
				RecipientCondition c = this._condition as RecipientCondition;
				return c.Regex;
			}
			else if (this._condition is SubjectCondition)
			{
				SubjectCondition c = this._condition as SubjectCondition;
				return c.Regex;
			}
			else if (this._condition is BodyCondition)
			{
				BodyCondition c = this._condition as BodyCondition;
				return c.Regex;
			}
			else
			{
				throw new Exception("Unknown type");
			}
		}

		private string GetUserText()
		{
			if (this._condition is SenderCondition)
			{
				SenderCondition c = this._condition as SenderCondition;
				return c.Sender;
			}
			else if (this._condition is RecipientCondition)
			{
				RecipientCondition c = this._condition as RecipientCondition;
				return c.Recipient;
			}
			else if (this._condition is SubjectCondition)
			{
				SubjectCondition c = this._condition as SubjectCondition;
				return c.Regex;
			}
			else if (this._condition is BodyCondition)
			{
				BodyCondition c = this._condition as BodyCondition;
				return c.Regex;
			}
			else
			{
				throw new Exception("Unknown type");
			}
		}

		private void Init()
		{
			this.radioRegex.SuspendLayout();
			this.radioEmailAddress.SuspendLayout();
			this.radioAddressBookEntry.SuspendLayout();

			this._regFont = this.nameTextBox.Font;
			this._boldFont = new Font(this._regFont, FontStyle.Bold|FontStyle.Underline);
			this.checkBoxNot.Checked = this._condition.Not;

			if (string.IsNullOrWhiteSpace(this.GetRegex()))
			{
				if (string.IsNullOrWhiteSpace(this.GetUserText()))
				{
					this._recipient = null;
					this.radioRegex.Checked = false;
					this.radioEmailAddress.Checked = true;
					this.radioAddressBookEntry.Checked = false;
				}
				else
				{
					this._recipient = this.GetRecipient();
					this.nameTextBox.Text = this.GetUserText();

					if (this._recipient.Resolve())
					{
						this.radioRegex.Checked = false;
						this.radioEmailAddress.Checked = false;
						this.radioAddressBookEntry.Checked = true;
					}
					else
					{
						this._recipient = null;

						this.radioRegex.Checked = false;
						this.radioEmailAddress.Checked = true;
						this.radioAddressBookEntry.Checked = false;
					}
				}
			}
			else
			{
				this.radioRegex.Checked = true;
				this.radioEmailAddress.Checked = false;
				this.radioAddressBookEntry.Checked = false;

			}

			this.radioRegex.ResumeLayout();
			this.radioEmailAddress.ResumeLayout();
			this.radioAddressBookEntry.ResumeLayout();

			this.ProcessRadioButtonChange();
		}

		private static object Clone(object o)
		{
			XmlSerializer ser = new XmlSerializer(o.GetType());
			MemoryStream stream = new MemoryStream();
			ser.Serialize(stream, o, new XmlSerializerNamespaces(new XmlQualifiedName[]{new XmlQualifiedName("")}));
			stream.Position = 0;
			Object newO = ser.Deserialize(stream);
			stream.Close();
			return newO;
		}

		private void ProcessRadioButtonChange()
		{
			if (this.radioRegex.Checked)
			{
				if (this.lastRadioButton == this.radioRegex)
				{
					return;
				}
				else
				{
					this.lastRadioButton = this.radioRegex;
				}

				if (this._recipient != null)
				{
					this.nameTextBox.Text = this._recipient.Address;
				}
				else
				{
				}

				this.buttonPickUser.Enabled = false;

				this.nameTextBox.ReadOnly = false;
				this.nameTextBox.Font = this._regFont;
			}
			else if (this.radioEmailAddress.Checked)
			{
				if (this.lastRadioButton == this.radioEmailAddress)
				{
					return;
				}
				else
				{
					this.lastRadioButton = this.radioEmailAddress;
				}

				if (this._recipient != null)
				{
					this.nameTextBox.Text = this._recipient.Address;
				}
				else
				{
				}

				this.buttonPickUser.Enabled = false;

				this.nameTextBox.ReadOnly = false;
				this.nameTextBox.Font = this._regFont;
			}
			else if (this.radioAddressBookEntry.Checked)
			{
				if (this.lastRadioButton == this.radioAddressBookEntry)
				{
					return;
				}
				else
				{
					this.lastRadioButton = this.radioAddressBookEntry;
				}

				if (this._recipient != null)
				{
					this.nameTextBox.Text = this._recipient.Name;
				}
				else
				{
					this.nameTextBox.Text = "";
				}

				this.buttonPickUser.Enabled = true;

				this.nameTextBox.ReadOnly = true;
				this.nameTextBox.Font = this._boldFont;
			}
			else
			{
				throw new Exception();
			}
		}

		private void radioAddressBookEntry_CheckedChanged(object sender, EventArgs e)
		{
			this.ProcessRadioButtonChange();
		}

		private void radioEmailAddress_CheckedChanged(object sender, EventArgs e)
		{
			this.ProcessRadioButtonChange();
		}

		private void radioRegex_CheckedChanged(object sender, EventArgs e)
		{
			this.ProcessRadioButtonChange();
		}

		private void buttonPickUser_Click(object sender, EventArgs e)
		{
			Outlook.SelectNamesDialog snd = this._application.Session.GetSelectNamesDialog();
			snd.SetDefaultDisplayMode(Outlook.OlDefaultSelectNamesDisplayMode.olDefaultSingleName);
			var result = snd.Display();

			if (result)
			{
				if (snd.Recipients.Count == 1)
				{
					// why is the first one index "1" and not index "0"?!?!
					this._recipient = snd.Recipients[1];

					this.ProcessRadioButtonChange();
				}
			}
		}

		private void cancelButton_Click(object sender, EventArgs e)
		{
			// nothing to do...
		}

		private void okButton_Click(object sender, EventArgs e)
		{
			if (this._condition is SenderCondition)
			{
				SenderCondition c = this._condition as SenderCondition;

				if (this.radioRegex.Checked)
				{
					c.Regex = this.nameTextBox.Text;
					c.Sender = null;
				}
				else if (this.radioEmailAddress.Checked)
				{
					c.Regex = null;
					c.Sender = this.nameTextBox.Text;
				}
				else if (this.radioAddressBookEntry.Checked)
				{
					if (this._recipient != null)
					{
						c.Regex = null;
						c.Sender = this._recipient.Address;
					}
					else
					{
						c.Regex = null;
						c.Sender = this.nameTextBox.Text;
					}
				}

			}
			else if (this._condition is RecipientCondition)
			{
				RecipientCondition c = this._condition as RecipientCondition;

				if (this.radioRegex.Checked)
				{
					c.Regex = this.nameTextBox.Text;
					c.Recipient = null;
				}
				else if (this.radioEmailAddress.Checked)
				{
					c.Regex = null;
					c.Recipient = this.nameTextBox.Text;
				}
				else if (this.radioAddressBookEntry.Checked)
				{
					if (this._recipient != null)
					{
						c.Regex = null;
						c.Recipient = this._recipient.Address;
					}
					else
					{
						c.Regex = null;
						c.Recipient = this.nameTextBox.Text;
					}
				}
			}
			else if (this._condition is SubjectCondition)
			{
				SubjectCondition c = this._condition as SubjectCondition;

				c.Regex = this.nameTextBox.Text;

			}
			else if (this._condition is BodyCondition)
			{
				BodyCondition c = this._condition as BodyCondition;

				c.Regex = this.nameTextBox.Text;
			}
			else
			{
				throw new Exception("Unknown type");
			}

			this._condition.Not = this.checkBoxNot.Checked;
		}

		private void ConditionDialog_Load(object sender, EventArgs e)
		{
		}

		private void ConditionDialog_FormClosed(object sender, FormClosedEventArgs e)
		{
			nop();
		}

		private void ConditionDialog_FormClosing(object sender, FormClosingEventArgs e)
		{
			nop();
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




