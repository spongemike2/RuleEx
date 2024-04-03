namespace RuleEx2025
{
	partial class ConditionDialog
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

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

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.okButton = new System.Windows.Forms.Button();
			this.cancelButton = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.radioRegex = new System.Windows.Forms.RadioButton();
			this.radioEmailAddress = new System.Windows.Forms.RadioButton();
			this.radioAddressBookEntry = new System.Windows.Forms.RadioButton();
			this.recipientNameLabel = new System.Windows.Forms.Label();
			this.nameTextBox = new System.Windows.Forms.TextBox();
			this.buttonPickUser = new System.Windows.Forms.Button();
			this.checkBoxNot = new System.Windows.Forms.CheckBox();
			this.groupBox1.SuspendLayout();
			this.SuspendLayout();
			// 
			// okButton
			// 
			this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.okButton.Location = new System.Drawing.Point(245, 142);
			this.okButton.Name = "okButton";
			this.okButton.Size = new System.Drawing.Size(75, 23);
			this.okButton.TabIndex = 0;
			this.okButton.Text = "OK";
			this.okButton.UseVisualStyleBackColor = true;
			this.okButton.Click += new System.EventHandler(this.okButton_Click);
			// 
			// cancelButton
			// 
			this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.cancelButton.Location = new System.Drawing.Point(245, 113);
			this.cancelButton.Name = "cancelButton";
			this.cancelButton.Size = new System.Drawing.Size(75, 23);
			this.cancelButton.TabIndex = 1;
			this.cancelButton.Text = "Cancel";
			this.cancelButton.UseVisualStyleBackColor = true;
			this.cancelButton.Click += new System.EventHandler(this.cancelButton_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.groupBox1.Controls.Add(this.radioRegex);
			this.groupBox1.Controls.Add(this.radioEmailAddress);
			this.groupBox1.Controls.Add(this.radioAddressBookEntry);
			this.groupBox1.Location = new System.Drawing.Point(13, 50);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(200, 115);
			this.groupBox1.TabIndex = 2;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Recipient";
			// 
			// radioRegex
			// 
			this.radioRegex.AutoSize = true;
			this.radioRegex.Location = new System.Drawing.Point(24, 83);
			this.radioRegex.Name = "radioRegex";
			this.radioRegex.Size = new System.Drawing.Size(116, 17);
			this.radioRegex.TabIndex = 2;
			this.radioRegex.TabStop = true;
			this.radioRegex.Text = "Regular Expression";
			this.radioRegex.UseVisualStyleBackColor = true;
			this.radioRegex.CheckedChanged += new System.EventHandler(this.radioRegex_CheckedChanged);
			// 
			// radioEmailAddress
			// 
			this.radioEmailAddress.AutoSize = true;
			this.radioEmailAddress.Location = new System.Drawing.Point(24, 52);
			this.radioEmailAddress.Name = "radioEmailAddress";
			this.radioEmailAddress.Size = new System.Drawing.Size(91, 17);
			this.radioEmailAddress.TabIndex = 1;
			this.radioEmailAddress.TabStop = true;
			this.radioEmailAddress.Text = "Email Address";
			this.radioEmailAddress.UseVisualStyleBackColor = true;
			this.radioEmailAddress.CheckedChanged += new System.EventHandler(this.radioEmailAddress_CheckedChanged);
			// 
			// radioAddressBookEntry
			// 
			this.radioAddressBookEntry.AutoSize = true;
			this.radioAddressBookEntry.Location = new System.Drawing.Point(24, 20);
			this.radioAddressBookEntry.Name = "radioAddressBookEntry";
			this.radioAddressBookEntry.Size = new System.Drawing.Size(118, 17);
			this.radioAddressBookEntry.TabIndex = 0;
			this.radioAddressBookEntry.TabStop = true;
			this.radioAddressBookEntry.Text = "Address Book Entry";
			this.radioAddressBookEntry.UseVisualStyleBackColor = true;
			this.radioAddressBookEntry.CheckedChanged += new System.EventHandler(this.radioAddressBookEntry_CheckedChanged);
			// 
			// recipientNameLabel
			// 
			this.recipientNameLabel.AutoSize = true;
			this.recipientNameLabel.Location = new System.Drawing.Point(13, 13);
			this.recipientNameLabel.Name = "recipientNameLabel";
			this.recipientNameLabel.Size = new System.Drawing.Size(31, 13);
			this.recipientNameLabel.TabIndex = 3;
			this.recipientNameLabel.Text = "Text:";
			// 
			// nameTextBox
			// 
			this.nameTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.nameTextBox.Location = new System.Drawing.Point(51, 9);
			this.nameTextBox.Name = "nameTextBox";
			this.nameTextBox.Size = new System.Drawing.Size(269, 20);
			this.nameTextBox.TabIndex = 4;
			// 
			// buttonPickUser
			// 
			this.buttonPickUser.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.buttonPickUser.Location = new System.Drawing.Point(245, 84);
			this.buttonPickUser.Name = "buttonPickUser";
			this.buttonPickUser.Size = new System.Drawing.Size(75, 23);
			this.buttonPickUser.TabIndex = 5;
			this.buttonPickUser.Text = "Pick User";
			this.buttonPickUser.UseVisualStyleBackColor = true;
			this.buttonPickUser.Click += new System.EventHandler(this.buttonPickUser_Click);
			// 
			// checkBoxNot
			// 
			this.checkBoxNot.AutoSize = true;
			this.checkBoxNot.Location = new System.Drawing.Point(245, 50);
			this.checkBoxNot.Name = "checkBoxNot";
			this.checkBoxNot.Size = new System.Drawing.Size(43, 17);
			this.checkBoxNot.TabIndex = 6;
			this.checkBoxNot.Text = "Not";
			this.checkBoxNot.TextAlign = System.Drawing.ContentAlignment.TopRight;
			this.checkBoxNot.UseVisualStyleBackColor = true;
			// 
			// ConditionDialog
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(332, 177);
			this.Controls.Add(this.checkBoxNot);
			this.Controls.Add(this.buttonPickUser);
			this.Controls.Add(this.nameTextBox);
			this.Controls.Add(this.recipientNameLabel);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.cancelButton);
			this.Controls.Add(this.okButton);
			this.Name = "ConditionDialog";
			this.Text = "Recipient Condition";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ConditionDialog_FormClosing);
			this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.ConditionDialog_FormClosed);
			this.Load += new System.EventHandler(this.ConditionDialog_Load);
			this.groupBox1.ResumeLayout(false);
			this.groupBox1.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button okButton;
		private System.Windows.Forms.Button cancelButton;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.RadioButton radioRegex;
		private System.Windows.Forms.RadioButton radioEmailAddress;
		private System.Windows.Forms.RadioButton radioAddressBookEntry;
		private System.Windows.Forms.Label recipientNameLabel;
		private System.Windows.Forms.TextBox nameTextBox;
		private System.Windows.Forms.Button buttonPickUser;
		private System.Windows.Forms.CheckBox checkBoxNot;
	}
}