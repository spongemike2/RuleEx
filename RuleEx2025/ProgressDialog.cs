using System;
using System.Windows.Forms;

namespace RuleEx2025
{
	public partial class ProgressDialog : Form
	{
		bool initialized = false;
		private double _progressPct;
		public double ProgressPct
		{
			get
			{
				return this._progressPct;
			}

			set
			{
				this._progressPct = value;

				if (this.initialized && this.Visible)
				{
					this.BeginInvoke(new System.Action(() =>
					{
						this.progressBar1.Value = (int)(this._progressPct * 1000.0);
					}));
				}
			}
		}


		private string _textBoxText;
		private string _oldTextBoxText;
		public string TextBoxText
		{
			get
			{
				return this._textBoxText;
			}

			set
			{
				this._textBoxText = value;

				if (_oldTextBoxText != _textBoxText)
				{
					_oldTextBoxText = _textBoxText;
					if (this.initialized && this.Visible)
					{
						this.BeginInvoke(new System.Action(() =>
						{
							this.textBox1.Text = this._textBoxText;
						}));
					}
				}
			}
		}

		private bool _wasCancelled = false;
		public bool WasCancelled
		{
			get
			{
				return this._wasCancelled;
			}
		}

		public ProgressDialog()
		{
			this.initialized = false;

			InitializeComponent();

			this.progressBar1.Visible = true;
			this.progressBar1.Minimum = 0;
			this.progressBar1.Maximum = 1000;
			this.progressBar1.Value = 0;
			this.progressBar1.Step = 1;
			this.progressBar1.Style = ProgressBarStyle.Blocks;//.Marquee;

			this._wasCancelled = false;
			this.initialized = true;
		}

		private void button1_Click(object sender, EventArgs e)
		{
			this._wasCancelled = true;
			this.Close();
		}
	}
}
