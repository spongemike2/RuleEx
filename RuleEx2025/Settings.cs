using System;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using Outlook = Microsoft.Office.Interop.Outlook;

//=====================================================================================================================================================================================================
//=====================================================================================================================================================================================================
namespace RuleEx2025
{
	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	public class Settings
	{
		[XmlAttribute] public bool Active;
		[XmlElement(typeof(Rule))]
		[XmlElement(typeof(BuildRule))]
		public Rule[] Rules
		{
			get
			{
				if (this._rules == null)
				{
					this._rules = new Rule[0];
				}

				return this._rules;
			}

			set
			{
				this._rules = value;
			}
		}

		private Rule[] _rules;

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public Settings()
		{
			this.Rules = new Rule[0];
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public Settings Clone()
		{
			XmlSerializer ser = new XmlSerializer(this.GetType());
			MemoryStream stream = new MemoryStream();
			ser.Serialize(stream, this, new XmlSerializerNamespaces(new XmlQualifiedName[]{new XmlQualifiedName("")}));
			stream.Position = 0;
			Settings settings = ser.Deserialize(stream) as Settings;
			stream.Close();
			return settings;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public static Settings Load(string fileName)
		{
			XmlSerializer ser = new XmlSerializer(typeof(Settings));
			using (TextReader reader = new StreamReader(fileName))
			{
				Settings settings = ser.Deserialize(reader) as Settings;
				reader.Close();
				ser = null;
				System.GC.Collect();
				settings.ReIndex();
				return settings;
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private static bool AreFilesIdentical(string path1, string path2)
		{
			using (FileStream file1 = new FileStream(path1, System.IO.FileMode.Open, System.IO.FileAccess.Read)) {
				using (FileStream file2 = new FileStream(path2, System.IO.FileMode.Open, System.IO.FileAccess.Read)) {

					if (file1.Length == file2.Length) {
						while (file1.Position < file1.Length) {
							if (file1.ReadByte() != file2.ReadByte())
							{
								return false;
							}
						}
						return true;
					}
					return false;
				}
			}

		}
		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void ReIndex()
		{
			int index = 0;
			foreach (var rule in this.Rules)
			{
				rule.Index = index;
				index += 100;
			}

			Array.Sort(this._rules, (Rule r1, Rule r2) => r1.Index - r2.Index);
		}


		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void Save(string fileName)
		{
			string backupFileName = null;
			string backupFilePath = null;

			// re-index all of the rules
			this.ReIndex();

			// back up any existing filename
			if (System.IO.File.Exists(fileName))
			{
				var info = new System.IO.FileInfo(fileName);

				string backupFolderName = "__RuleExBackup";
				string backupFolderPath = System.IO.Path.Combine(info.DirectoryName, backupFolderName);

				// ensure that the backup folder exists
				Directory.CreateDirectory(backupFolderPath);

				int c = 0;
				backupFileName = string.Format(@"{0}_{1:0000}{2}", System.IO.Path.GetFileNameWithoutExtension(info.Name),  c, info.Extension);
				backupFilePath = System.IO.Path.Combine(backupFolderPath, backupFileName);

				while (System.IO.File.Exists(backupFilePath))
				{
					++c;
					backupFileName = string.Format(@"{0}_{1:0000}{2}", System.IO.Path.GetFileNameWithoutExtension(info.Name), c, info.Extension);
					backupFilePath = System.IO.Path.Combine(backupFolderPath, backupFileName);
				}

				System.IO.Directory.CreateDirectory(backupFolderPath);

				if (System.IO.File.Exists(fileName))
				{
                    System.IO.File.Move(fileName, backupFilePath);
                }
            }

			XmlSerializer ser = new XmlSerializer(this.GetType());
			TextWriter writer = new StreamWriter(fileName, false, System.Text.Encoding.ASCII);
			ser.Serialize(writer, this, new XmlSerializerNamespaces(new XmlQualifiedName[]{new XmlQualifiedName("")}));
			writer.Close();

			if (backupFilePath != null)
			{
				// compare the file with the backup... if they are the same, then no need to save the backup
				if (AreFilesIdentical(backupFilePath, fileName))
				{
					// delete the backup
					System.IO.File.Delete(backupFilePath);
				}
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public void Resolve(Outlook.Application application)
		{
			foreach (Rule rule in this.Rules)
			{
				if (rule.Actions != null)
				{
					foreach (Action action in rule.Actions)
					{
						if (action is MoveAction)
						{
							MoveAction moveAction = action as MoveAction;
							try
							{
								moveAction.Folder = application.Session.GetFolderFromID(moveAction.FolderName);
								moveAction.FolderPath = moveAction.Folder.FullFolderPath;
							}
							catch (Exception ex)
							{
								this.nop(ex);
								rule.Active = false;
							}
						}
					}
				}
			}
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private void nop(object o=null)
		{
		}
	}
}
