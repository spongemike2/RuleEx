namespace RuleEx2025
{
	partial class RuleDialog
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
			this.nameLabel = new System.Windows.Forms.Label();
			this.lastRunLabel = new System.Windows.Forms.Label();
			this.runCountLabel = new System.Windows.Forms.Label();
			this.indexLabel = new System.Windows.Forms.Label();
			this.nameTextBox = new System.Windows.Forms.TextBox();
			this.lastRunTextBox = new System.Windows.Forms.TextBox();
			this.runCountTextBox = new System.Windows.Forms.TextBox();
			this.indexTextBox = new System.Windows.Forms.TextBox();
			this.activeCheckBox = new System.Windows.Forms.CheckBox();
			this.finalCheckBox = new System.Windows.Forms.CheckBox();
			this.operatorGroup = new System.Windows.Forms.GroupBox();
			this.andOperatorButton = new System.Windows.Forms.RadioButton();
			this.orOperatorButton = new System.Windows.Forms.RadioButton();
			this.operatorGroup.SuspendLayout();
			this.SuspendLayout();
			// 
			// okButton
			// 
			this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.okButton.DialogResult = System.Windows.Forms.DialogResult.OK;
			this.okButton.Location = new System.Drawing.Point(241, 195);
			this.okButton.Name = "okButton";
			this.okButton.Size = new System.Drawing.Size(75, 23);
			this.okButton.TabIndex = 1;
			this.okButton.Text = "OK";
			this.okButton.UseVisualStyleBackColor = true;
			this.okButton.Click += new System.EventHandler(this.okButton_Click);
			// 
			// cancelButton
			// 
			this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.cancelButton.Location = new System.Drawing.Point(322, 195);
			this.cancelButton.Name = "cancelButton";
			this.cancelButton.Size = new System.Drawing.Size(75, 23);
			this.cancelButton.TabIndex = 2;
			this.cancelButton.Text = "Cancel";
			this.cancelButton.UseVisualStyleBackColor = true;
			// 
			// nameLabel
			// 
			this.nameLabel.AutoSize = true;
			this.nameLabel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.nameLabel.Location = new System.Drawing.Point(10, 16);
			this.nameLabel.Name = "nameLabel";
			this.nameLabel.Size = new System.Drawing.Size(63, 13);
			this.nameLabel.TabIndex = 3;
			this.nameLabel.Text = "Rule Name:";
			// 
			// lastRunLabel
			// 
			this.lastRunLabel.AutoSize = true;
			this.lastRunLabel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.lastRunLabel.Location = new System.Drawing.Point(20, 46);
			this.lastRunLabel.Name = "lastRunLabel";
			this.lastRunLabel.Size = new System.Drawing.Size(53, 13);
			this.lastRunLabel.TabIndex = 4;
			this.lastRunLabel.Text = "Last Run:";
			// 
			// runCountLabel
			// 
			this.runCountLabel.AutoSize = true;
			this.runCountLabel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.runCountLabel.Location = new System.Drawing.Point(12, 76);
			this.runCountLabel.Name = "runCountLabel";
			this.runCountLabel.Size = new System.Drawing.Size(61, 13);
			this.runCountLabel.TabIndex = 5;
			this.runCountLabel.Text = "Run Count:";
			//
			// indexLabel
			//
			this.indexLabel.AutoSize = true;
			this.indexLabel.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
			this.indexLabel.Location = new System.Drawing.Point(12, 106);
			this.indexLabel.Name = "indexLabel";
			this.indexLabel.Size = new System.Drawing.Size(61, 13);
			this.indexLabel.TabIndex = 5;
			this.indexLabel.Text = "Index:";
			// 
			// nameTextBox
			// 
			this.nameTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.nameTextBox.Location = new System.Drawing.Point(79, 13);
			this.nameTextBox.Name = "nameTextBox";
			this.nameTextBox.Size = new System.Drawing.Size(318, 20);
			this.nameTextBox.TabIndex = 6;
			// 
			// lastRunTextBox
			// 
			this.lastRunTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.lastRunTextBox.Location = new System.Drawing.Point(79, 43);
			this.lastRunTextBox.Name = "lastRunTextBox";
			this.lastRunTextBox.ReadOnly = true;
			this.lastRunTextBox.Size = new System.Drawing.Size(318, 20);
			this.lastRunTextBox.TabIndex = 7;
			// 
			// runCountTextBox
			// 
			this.runCountTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this.runCountTextBox.Location = new System.Drawing.Point(79, 73);
			this.runCountTextBox.Name = "runCountTextBox";
			this.runCountTextBox.ReadOnly = true;
			this.runCountTextBox.Size = new System.Drawing.Size(318, 20);
			this.runCountTextBox.TabIndex = 8;
			//
			// indexTextBox
			//
			this.indexTextBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
			| System.Windows.Forms.AnchorStyles.Right)));
			this.indexTextBox.Location = new System.Drawing.Point(79, 103);
			this.indexTextBox.Name = "indexTextBox";
			this.indexTextBox.ReadOnly = true;
			this.indexTextBox.Size = new System.Drawing.Size(318, 20);
			this.indexTextBox.TabIndex = 8;
			// 
			// activeCheckBox
			// 
			this.activeCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.activeCheckBox.AutoSize = true;
			this.activeCheckBox.Location = new System.Drawing.Point(19, 166);
			this.activeCheckBox.Name = "activeCheckBox";
			this.activeCheckBox.Size = new System.Drawing.Size(56, 17);
			this.activeCheckBox.TabIndex = 9;
			this.activeCheckBox.Text = "Active";
			this.activeCheckBox.UseVisualStyleBackColor = true;
			// 
			// finalCheckBox
			// 
			this.finalCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.finalCheckBox.AutoSize = true;
			this.finalCheckBox.Location = new System.Drawing.Point(19, 193);
			this.finalCheckBox.Name = "finalCheckBox";
			this.finalCheckBox.Size = new System.Drawing.Size(48, 17);
			this.finalCheckBox.TabIndex = 10;
			this.finalCheckBox.Text = "Final";
			this.finalCheckBox.UseVisualStyleBackColor = true;
			// 
			// operatorGroup
			// 
			this.operatorGroup.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
			this.operatorGroup.Controls.Add(this.andOperatorButton);
			this.operatorGroup.Controls.Add(this.orOperatorButton);
			this.operatorGroup.Location = new System.Drawing.Point(103, 153);
			this.operatorGroup.Name = "operatorGroup";
			this.operatorGroup.Size = new System.Drawing.Size(93, 65);
			this.operatorGroup.TabIndex = 11;
			this.operatorGroup.TabStop = false;
			this.operatorGroup.Text = "Operator";
			this.operatorGroup.Enter += new System.EventHandler(this.operatorGroup_Enter);
			// 
			// andOperatorButton
			// 
			this.andOperatorButton.AutoSize = true;
			this.andOperatorButton.Location = new System.Drawing.Point(7, 42);
			this.andOperatorButton.Name = "andOperatorButton";
			this.andOperatorButton.Size = new System.Drawing.Size(44, 17);
			this.andOperatorButton.TabIndex = 1;
			this.andOperatorButton.TabStop = true;
			this.andOperatorButton.Text = "And";
			this.andOperatorButton.UseVisualStyleBackColor = true;
			// 
			// orOperatorButton
			// 
			this.orOperatorButton.AutoSize = true;
			this.orOperatorButton.Location = new System.Drawing.Point(7, 20);
			this.orOperatorButton.Name = "orOperatorButton";
			this.orOperatorButton.Size = new System.Drawing.Size(36, 17);
			this.orOperatorButton.TabIndex = 0;
			this.orOperatorButton.TabStop = true;
			this.orOperatorButton.Text = "Or";
			this.orOperatorButton.UseVisualStyleBackColor = true;
			// 
			// RuleDialog
			// 
			this.AcceptButton = this.okButton;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.CancelButton = this.cancelButton;
			this.ClientSize = new System.Drawing.Size(409, 230);
			this.Controls.Add(this.operatorGroup);
			this.Controls.Add(this.finalCheckBox);
			this.Controls.Add(this.activeCheckBox);
			this.Controls.Add(this.runCountTextBox);
			this.Controls.Add(this.indexTextBox);
			this.Controls.Add(this.lastRunTextBox);
			this.Controls.Add(this.nameTextBox);
			this.Controls.Add(this.runCountLabel);
			this.Controls.Add(this.indexLabel);
			this.Controls.Add(this.lastRunLabel);
			this.Controls.Add(this.nameLabel);
			this.Controls.Add(this.cancelButton);
			this.Controls.Add(this.okButton);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow;
			this.Name = "RuleDialog";
			this.Text = "Outlook Rule";
			this.Load += new System.EventHandler(this.RuleDialog_Load);
			this.operatorGroup.ResumeLayout(false);
			this.operatorGroup.PerformLayout();
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Button okButton;
		private System.Windows.Forms.Button cancelButton;
		private System.Windows.Forms.Label nameLabel;
		private System.Windows.Forms.Label lastRunLabel;
		private System.Windows.Forms.Label runCountLabel;
		private System.Windows.Forms.Label indexLabel;
		private System.Windows.Forms.TextBox nameTextBox;
		private System.Windows.Forms.TextBox lastRunTextBox;
		private System.Windows.Forms.TextBox runCountTextBox;
		private System.Windows.Forms.TextBox indexTextBox;
		private System.Windows.Forms.CheckBox activeCheckBox;
		private System.Windows.Forms.CheckBox finalCheckBox;
		private System.Windows.Forms.GroupBox operatorGroup;
		private System.Windows.Forms.RadioButton andOperatorButton;
		private System.Windows.Forms.RadioButton orOperatorButton;
	}
}