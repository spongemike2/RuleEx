using System;
using System.IO;

//=====================================================================================================================================================================================================
//=====================================================================================================================================================================================================
namespace RuleEx2025
{
	public abstract class Logger
	{
		public abstract void Write(string value);
		public abstract void WriteLine(string value);
		public abstract void WriteFormat(string value, params object[] o);
		public abstract void WriteLineFormat(string value, params object[] o);
		static private Logger _the;

		static public Logger the
		{
			get
			{
				return _the;
			}

			set
			{
				_the = value;
			}
		}
	}

	public class FileLogger : Logger
	{
		private string _FileName;
		private StreamWriter	_sw;

		public FileLogger()
		{
			this._FileName = Path.Combine(Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData), "RuleEx2025.log");
			Logger.the = this;
		}

		public FileLogger(string fileName)
		{
			this._FileName = fileName;
			Logger.the = this;
		}

		~FileLogger()
		{
		}

		private void _Write(string value)
		{
			if (this._sw == null)
			{
				if (File.Exists(this._FileName))
				{
					this._sw = File.AppendText(this._FileName);
				}
				else
				{
					this._sw = File.CreateText(this._FileName);
				}

				this._sw.WriteLine("========================================================================================================================================================================================================");
				this._sw.WriteLine(string.Format("RuleEx2025 started on {0:yyyy-MM-dd:hh:mm:sstt}", DateTime.Now));
			}

			this._sw.Write(value);
			this._sw.Flush();
			System.Diagnostics.Debug.Write(value);
		}

		public override void Write(string value)
		{
			this._Write(value);
		}

		public override void WriteLine(string value)
		{
			this._Write(value);
			this._Write("\n");
		}

		public override void WriteFormat(string value, params object[] o)
		{
			this._Write(string.Format(value, o));
		}
		public override void WriteLineFormat(string value, params object[] o)
		{
			this._Write(string.Format(value, o));
			this._Write("\n");
		}
	}
}
