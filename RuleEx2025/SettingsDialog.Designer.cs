namespace RuleEx2025
{
	partial class SettingsDialog
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
			this._cancelButton = new System.Windows.Forms.Button();
			this._saveButton = new System.Windows.Forms.Button();
			this._settingsTreeView = new System.Windows.Forms.TreeView();
			this.SuspendLayout();
			// 
			// _cancelButton
			// 
			this._cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this._cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this._cancelButton.Location = new System.Drawing.Point(410, 665);
			this._cancelButton.Name = "_cancelButton";
			this._cancelButton.Size = new System.Drawing.Size(75, 20);
			this._cancelButton.TabIndex = 0;
			this._cancelButton.Text = "Cancel";
			this._cancelButton.UseVisualStyleBackColor = true;
			this._cancelButton.Click += new System.EventHandler(this._cancelButton_Click);
			// 
			// _saveButton
			// 
			this._saveButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
			this._saveButton.DialogResult = System.Windows.Forms.DialogResult.OK;
			this._saveButton.Location = new System.Drawing.Point(320, 665);
			this._saveButton.Name = "_saveButton";
			this._saveButton.Size = new System.Drawing.Size(75, 20);
			this._saveButton.TabIndex = 1;
			this._saveButton.Text = "Save";
			this._saveButton.UseVisualStyleBackColor = true;
			this._saveButton.Click += new System.EventHandler(this._saveButton_Click);
			// 
			// _settingsTreeView
			// 
			this._settingsTreeView.AllowDrop = true;
			this._settingsTreeView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
			this._settingsTreeView.Location = new System.Drawing.Point(15, 15);
			this._settingsTreeView.Name = "_settingsTreeView";
			this._settingsTreeView.Size = new System.Drawing.Size(470, 635);
			this._settingsTreeView.TabIndex = 2;
			this._settingsTreeView.BeforeExpand += new System.Windows.Forms.TreeViewCancelEventHandler(this._settingsTreeView_BeforeExpand);
			this._settingsTreeView.AfterExpand += new System.Windows.Forms.TreeViewEventHandler(this._settingsTreeView_AfterExpand);
			this._settingsTreeView.ItemDrag += new System.Windows.Forms.ItemDragEventHandler(this._settingsTreeView_ItemDrag);
			this._settingsTreeView.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this._settingsTreeView_NodeMouseClick);
			this._settingsTreeView.NodeMouseDoubleClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this._settingsTreeView_NodeMouseDoubleClick);
			this._settingsTreeView.Click += new System.EventHandler(this._settingsTreeView_Click);
			this._settingsTreeView.DragDrop += new System.Windows.Forms.DragEventHandler(this._settingsTreeView_DragDrop);
			this._settingsTreeView.DragEnter += new System.Windows.Forms.DragEventHandler(this._settingsTreeView_DragEnter);
			this._settingsTreeView.DragOver += new System.Windows.Forms.DragEventHandler(this._settingsTreeView_DragOver);
			this._settingsTreeView.DragLeave += new System.EventHandler(this._settingsTreeView_DragLeave);
			this._settingsTreeView.GiveFeedback += new System.Windows.Forms.GiveFeedbackEventHandler(this._settingsTreeView_GiveFeedback);
			this._settingsTreeView.QueryContinueDrag += new System.Windows.Forms.QueryContinueDragEventHandler(this._settingsTreeView_QueryContinueDrag);
			this._settingsTreeView.DoubleClick += new System.EventHandler(this._settingsTreeView_DoubleClick);
			this._settingsTreeView.MouseClick += new System.Windows.Forms.MouseEventHandler(this._settingsTreeView_MouseClick);
			this._settingsTreeView.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this._settingsTreeView_MouseDoubleClick);
			// 
			// SettingsDialog
			// 
			this.AcceptButton = this._saveButton;
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(500, 700);
			this.Controls.Add(this._settingsTreeView);
			this.Controls.Add(this._saveButton);
			this.Controls.Add(this._cancelButton);
			this.Name = "SettingsDialog";
			this.Text = "RuleEx 2025 Settings";
			this.Load += new System.EventHandler(this.SettingsDialog_Load);
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.Button _cancelButton;
		private System.Windows.Forms.Button _saveButton;
		private System.Windows.Forms.TreeView _settingsTreeView;
	}
}