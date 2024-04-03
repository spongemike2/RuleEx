using System;
using System.Collections;
using System.Collections.Generic;
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
	public class Rule : ConditionGroup
	{
		[XmlAttribute] public string Name;
		[XmlAttribute] public bool Active;
		[XmlAttribute] public bool Final;
		[XmlAttribute] public int Index;
		[XmlAttribute] public DateTime LastRun;
		[XmlAttribute] public long RunCount;

		[XmlElement(typeof(MoveAction))]
		[XmlElement(typeof(MarkAsReadAction))]
		[XmlElement(typeof(BuildMoveAction))]
		[XmlElement(typeof(TagCategoryAction))]
		[XmlElement(typeof(ForwardAction))]
		[XmlElement(typeof(FixLinksAction))]
		public Action[]	Actions;

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public new Rule Clone()
		{
			XmlSerializer ser = new XmlSerializer(this.GetType());
			MemoryStream stream = new MemoryStream();
			ser.Serialize(stream, this, new XmlSerializerNamespaces(new XmlQualifiedName[]{new XmlQualifiedName("")}));
			stream.Position = 0;
			Rule rule = ser.Deserialize(stream) as Rule;
			stream.Close();
			return rule;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public static Rule CreateBuildMoveRule_old(int index, string name, bool active, string parentFolder, string[] buildAccounts, string[] buildGroups)
		{
			BuildRule rule = new BuildRule();

			rule.Index = index;
			rule.Name = name;
			rule.Active = active;
			rule.Final = true;
			rule.Operator = ConditionGroup.GroupingOperator.Or;

			ArrayList conditions = new ArrayList(buildGroups.Length);

			foreach (string buildGroup in buildGroups)
			{
				conditions.Add(new RecipientCondition(buildGroup, null));
			}

			rule.Actions = new Action[] { new BuildMoveAction(parentFolder, buildGroups, buildAccounts) };
			rule.Conditions = conditions.ToArray(typeof(Condition)) as Condition[];
			return rule;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public Rule()
		{
			this.Final = true;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public Rule(int index, string name, bool active, ConditionGroup conditions, Action[] actions)
		{
			this.Index = index;
			this.Name = name;
			this.Active = active;
			this.Actions = actions;
			this.Final = true;

			this.Operator = conditions.Operator;
			this.Conditions = conditions.Conditions;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public Rule(int index, string name, bool active, Condition conditions, Action[] actions)
		{
			this.Index = index;
			this.Name = name;
			this.Active = active;
			this.Actions = actions;
			this.Final = true;

			this.Operator = ConditionGroup.GroupingOperator.Or;
			this.Conditions = new Condition[] { conditions };
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public Rule(int index, string name, bool active, Condition[] conditions, Action[] actions)
		{
			this.Index = index;
			this.Name = name;
			this.Active = active;
			this.Actions = actions;
			this.Final = true;

			this.Operator = ConditionGroup.GroupingOperator.Or;
			this.Conditions = conditions;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public Rule(int index, string name, bool active, Action[] actions, params Condition[] conditions)
		{
			this.Index = index;
			this.Name = name;
			this.Active = active;
			this.Actions = actions;
			this.Final = true;

			this.Operator = ConditionGroup.GroupingOperator.Or;
			this.Conditions = conditions;
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public virtual void InvokeActions(dynamic mailItem)
		{
			if (mailItem is Outlook.MailItem)
			{
				Outlook.MailItem item = mailItem as Outlook.MailItem;
				Logger.the.WriteLine(string.Format("Ran rule {1,4} at {2:yyyy-MM-dd:hh:mm:sstt}: Name \"{0}\" Subject: \"{3}\"", this.Name, this.Index, DateTime.Now, item.Subject));
			}
			else if (mailItem is Outlook.MeetingItem)
			{
				Outlook.MeetingItem item = mailItem as Outlook.MeetingItem;
				Logger.the.WriteLine(string.Format("Ran rule {1,4} at {2:yyyy-MM-dd:hh:mm:sstt}: Name \"{0}\" Subject: \"{3}\"", this.Name, this.Index, DateTime.Now, item.Subject));
			}
			else
			{
				Logger.the.WriteLine(string.Format("Ran rule {1,4} at {2:yyyy-MM-dd:hh:mm:sstt}: Name \"{0}\" Subject: Unknown", this.Name, this.Index, DateTime.Now));
			}

			++this.RunCount;
			this.LastRun = DateTime.Now;

			foreach (Action action in this.Actions)
			{
				mailItem = action.Invoke(this, mailItem);
			}
		}
	}

	//=================================================================================================================================================================================================
	//=================================================================================================================================================================================================
	public class BuildRule : Rule
	{
		[XmlElement("Recipient")] public string[] DistributionLists;
		[XmlElement("BuildAccount")] public string[] BuildAccounts;
		[XmlAttribute] public string ParentFolder;
		[XmlIgnore] public bool _initialized = false;
		[XmlIgnore] private BuildMoveAction _action;
		[XmlIgnore] private ConditionGroup _conditions;

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public BuildRule()
		{
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public BuildRule(int index, string name, bool active, string parentFolder, string[] buildAccounts, string[] buildGroups)
		{
			this.Index = index;
			this.Name = name;
			this.Active = active;
			this.Final = true;
			this.Operator = ConditionGroup.GroupingOperator.Or;

			this.DistributionLists = buildGroups;
			this.BuildAccounts = buildAccounts;
			this.ParentFolder = parentFolder;

			this.Initialize();
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		private void Initialize()
		{
			this._initialized = true;

			if (this.DistributionLists == null)
			{
				this.DistributionLists = new string[0];
			}

			RecipientCondition[] conditions = new RecipientCondition[this.DistributionLists.Length];

			for(int i=0; i<this.DistributionLists.Length; ++i)
			{
				conditions[i] = new RecipientCondition(this.DistributionLists[i], null);
			}

			this._conditions = new ConditionGroup(ConditionGroup.GroupingOperator.Or, conditions);
			this._action = new BuildMoveAction(this.ParentFolder, this.DistributionLists, this.BuildAccounts);

			this.Actions = new Action[] {};
			this.Conditions = new Condition[] {};
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		protected override bool _Test(dynamic mailItem, Dictionary<string,string> meta)
		{
			if (!this._initialized)
			{
				this.Initialize();
			}

			return this._conditions.Test(mailItem, meta);
		}

		//=============================================================================================================================================================================================
		//=============================================================================================================================================================================================
		public override void InvokeActions(dynamic mailItem)
		{
			base.InvokeActions((object)mailItem);
			mailItem = this._action.Invoke(this, mailItem);
		}

	}
}


