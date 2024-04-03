using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Serialization;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RuleEx2025
{
	public partial class SettingsDialog : Form
	{
		private Settings _settings;
		private Outlook.Application _application;

		private const int dW = 500; // dialog windth
		private const int dH = 700; // dialog height
		private const int bW = 75; // button width
		private const int bH = 20; // button height
		private const int S = 15; // spacer

		public Settings Settings
		{
			get
			{
				return this._settings;
			}
		}

		private void nop(object o=null)
		{
		}

		public SettingsDialog()
		{
			InitializeComponent();
		}

		public SettingsDialog(Settings settings, Outlook.Application application, int initialRule = 0) : this()
		{
			//
			// clone the settings... for now
			//
			this._settings = settings.Clone();
			this._application = application;

			//
			// new tree view
			//
			if (true)
			{
				this._settingsTreeView.SuspendLayout();
				this._settingsTreeView.BeginUpdate();
				TreeNode rulesNode = this._settingsTreeView.Nodes.Add("");
				rulesNode.Collapse();
				rulesNode.Tag = this._settings.Rules;
				SetNodeText(rulesNode);

				int i = 0;
				int c = this._settings.Rules.Length;

				foreach (var rule in this._settings.Rules)
				{
					++i;

					TreeNode node = rulesNode.Nodes.Add("");
					node.ToolTipText = "Edit the rule";
					node.Tag = rule;
					SetNodeText(node);

					AddConditions(node);

					TreeNode actionsNode = node.Nodes.Add("");
					if (rule.Actions == null)
					{
						rule.Actions = new Action[0];
					}
					actionsNode.Tag = rule.Actions;
					SetNodeText(actionsNode);

					foreach (var action in rule.Actions)
					{
						TreeNode actionNode = actionsNode.Nodes.Add("");
						actionNode.Tag = action;
						SetNodeText(actionNode);
					}
				}

				this._settingsTreeView.Click += _settingsTreeView_Click;
				this._settingsTreeView.DoubleClick += _settingsTreeView_DoubleClick;

				if (initialRule >=0 && initialRule < this._settingsTreeView.Nodes[0].Nodes.Count)
				{
					this._settingsTreeView.SelectedNode = this._settingsTreeView.Nodes[0].Nodes[initialRule];
					this._settingsTreeView.Nodes[0].Nodes[initialRule].ExpandAll();
				}
				else
				{
				}


				//this._settingsTreeView.ContextMenu.

				this._settingsTreeView.EndUpdate();
				this._settingsTreeView.ResumeLayout();
			}
		}

		//=========================================================================================
		//=========================================================================================
		private void SetNodeTextRecursive(TreeNode node)
		{
			SetNodeText(node);

			foreach (TreeNode childNode in node.Nodes)
			{
				SetNodeTextRecursive(childNode);
			}
		}

		//=========================================================================================
		//=========================================================================================
		private void SetNodeText(TreeNode node)
		{
			if (node.Tag is Rule[])
			{
				node.Text = "Rules";
			}
			else if (node.Tag is Rule)
			{
				node.Text = ((Rule)node.Tag).Name;
			}
			else if (node.Tag is Condition[])
			{
				ConditionGroup.GroupingOperator op;

				if (node.Parent.Tag is Rule)
				{
					op = ((Rule)node.Parent.Tag).Operator;
				}
				else if (node.Parent.Tag is ConditionGroup)
				{
					op = ((ConditionGroup)node.Parent.Tag).Operator;
				}
				else
				{
					throw new Exception("Error");
				}

				string conditionGroupName = string.Format("Conditions ({0})", op == ConditionGroup.GroupingOperator.Or ? "OR" : "AND");
				node.Text = conditionGroupName;
			}
			else if (node.Tag is Condition)
			{
				Condition condition = node.Tag as Condition;

				if (condition is RecipientCondition)
				{
					RecipientCondition c = condition as RecipientCondition;

					if (string.IsNullOrEmpty(c.Regex))
					{
						//
						// see if it's expanded
						//
						bool allExpanded = true;

						TreeNode p = node.Parent;
						while (p != null)
						{
							if (!p.IsExpanded)
							{
								allExpanded = false;
								break;
							}

							p = p.Parent;
						}

						if (allExpanded)
						{
							Microsoft.Office.Interop.Outlook.Recipient r = this._application.Session.CreateRecipient(c.Recipient);

							if (r.Resolve())
							{
								node.Text = string.Format("Recipient {0}: {1}", (c.Not ? "is NOT" : "is"), r.Name);
							}
							else
							{
								string recipientAddress = c.Recipient;

								if (recipientAddress.Length > 40)
								{
									recipientAddress = recipientAddress.Substring(recipientAddress.Length-40,40);
								}

								node.Text = string.Format("Recipient {0}: {1}", (c.Not ? "is NOT" : "is"), recipientAddress);
							}
						}
						else
						{
							node.Text = string.Format("Unknown");
						}
					}
					else
					{
						node.Text = string.Format("Recipient {0} (Regex): {1}", (c.Not ? "is NOT" : "is"), c.Regex);
					}
				}
				else if (condition is SenderCondition)
				{
					SenderCondition c = condition as SenderCondition;

					if (string.IsNullOrEmpty(c.Regex))
					{
						//
						// see if it's expanded
						//
						bool allExpanded = true;

						TreeNode p = node.Parent;
						while (p != null)
						{
							if (!p.IsExpanded)
							{
								allExpanded = false;
								break;
							}

							p = p.Parent;
						}

						if (allExpanded)
						{
							Microsoft.Office.Interop.Outlook.Recipient r = this._application.Session.CreateRecipient(c.Sender);

							if (r.Resolve())
							{
								node.Text = string.Format("Sender {0}: {1}", (c.Not ? "is NOT" : "is"), r.Name);
							}
							else
							{
								string senderAddress = c.Sender;

								if (senderAddress.Length > 40)
								{
									senderAddress = senderAddress.Substring(senderAddress.Length-40,40);
								}

								node.Text = string.Format("Sender {0}: {1}", (c.Not ? "is NOT" : "is"), senderAddress);
							}
						}
						else
						{
							node.Text = string.Format("Unknown");
						}
					}
					else
					{
						node.Text = string.Format("Sender {0} (Regex): {1}", (c.Not ? "is NOT" : "is"), c.Regex);
					}

				}
				else if (condition is SubjectCondition)
				{
					SubjectCondition c = condition as SubjectCondition;
					node.Text = string.Format("Subject Condition{1}: {0}", c.Regex, (c.Not ? " (NOT)" : ""));
				}
				else if (condition is BodyCondition)
				{
					BodyCondition c = condition as BodyCondition;
					node.Text = string.Format("Body Condition{1}: {0}", c.Regex, (c.Not ? " (NOT)" : ""));
				}
				else if (condition is ConditionGroup)
				{
					ConditionGroup c = condition as ConditionGroup;
					node.Text = string.Format("Group ({0})", c.Operator == ConditionGroup.GroupingOperator.Or ? "OR" : "AND");
				}
				else
				{
					throw new Exception(string.Format("Unknown condition type: {0}", condition.GetType().FullName));
				}
			}
			else if (node.Tag is Action[])
			{
				node.Text = "Actions";
			}
			else if (node.Tag is Action)
			{
				Action action = (Action)node.Tag;
				string actionName;

				if (action is BuildMoveAction)
				{
					BuildMoveAction a = action as BuildMoveAction;
					actionName = string.Format("Build Move: {0}", a.ParentFolder.FolderPath);
				}
				else if (action is MoveAction)
				{
					MoveAction a = action as MoveAction;
					actionName = string.Format("Move: {0}", a.FolderPath);
				}
				else if (action is MarkAsReadAction)
				{
					MarkAsReadAction a = action as MarkAsReadAction;
					actionName = string.Format("Mark as Read");
				}
				else if (action is TagCategoryAction)
				{
					TagCategoryAction a = action as TagCategoryAction;
					actionName = string.Format("Tag Category: {0}", a.Category);
				}
				else if (action is ForwardAction)
				{
					ForwardAction a = action as ForwardAction;
					actionName = string.Format("Forward: {0}", a.EmailAddress);
				}
				else if (action is FixLinksAction)
				{
					FixLinksAction a = action as FixLinksAction;
					actionName = string.Format("Fix Links");
				}
				else
				{
					throw new Exception(string.Format("Unknown action type: {0}", action.GetType().FullName));
				}

				node.Text = actionName;
			}
			else
			{
				throw new Exception();
			}
		}

		//=========================================================================================
		//=========================================================================================
		private void _saveButton_Click(object sender, EventArgs e)
		{
			this._settings = this._settings.Clone();
		}

		//=========================================================================================
		//=========================================================================================
		private void _cancelButton_Click(object sender, EventArgs e)
		{

		}

		//=========================================================================================
		//=========================================================================================
		private void _settingsTreeView_Click(object sender, EventArgs e)
		{

		}

		//=========================================================================================
		//=========================================================================================
		private void _settingsTreeView_DoubleClick(object sender, EventArgs e)
		{

		}

		//=========================================================================================
		// Inserts the given condition into the tree view before this node
		//=========================================================================================
		private bool InsertConditionBeforeNode(TreeNode node, Condition condition)
		{
			if (condition == null)
			{
				return false;
			}

			if (node == null)
			{
				return false;
			}

			if (node.Tag is Rule)
			{
				// not valid
				return false;
			}

			if (node.Tag is Condition)
			{
				Condition targetCondition = node.Tag as Condition;
				Condition[] targetConditions = node.Parent.Tag as Condition[];

				if (targetCondition == null)
				{
					throw new Exception();
				}

				if (targetConditions == null)
				{
					throw new Exception();
				}

				// duplicate the condition list
				List<Condition> list = new List<Condition>();
				foreach (Condition c in targetConditions)
				{
					list.Add(c);
				}

				// find the index of the target condition
				int index = 0;
				for (; index < targetConditions.Length; ++index)
				{
					if (targetConditions[index].Id == targetCondition.Id)
					{
						break;
					}
				}

				// this should never happen
				if (index >= targetConditions.Length)
				{
					throw new Exception();
				}


				Condition newCondition = condition.Clone();
				TreeNode newNode = node.Parent.Nodes.Insert(index, "");
				newNode.Tag = newCondition;
				list.Insert(index, newCondition);

				// now, set the parent's conditions
				if (node.Parent.Parent.Tag is Rule)
				{
					Rule rule = node.Parent.Parent.Tag as Rule;
					rule.Conditions = list.ToArray();
					node.Parent.Tag = rule.Conditions;
				}
				else if (node.Parent.Parent.Tag is ConditionGroup)
				{
					ConditionGroup cg = node.Parent.Parent.Tag as ConditionGroup;
					cg.Conditions = list.ToArray();
					node.Parent.Tag = cg.Conditions;
				}
				else
				{
					throw new Exception();
				}

				SetNodeText(newNode);
				return true;
			}
			else
			{
				return false;
			}
		}

		private bool DeleteConditionOrActionAtNode(TreeNode node)
		{
			if (node == null)
			{
				return false;
			}

			if (node.Tag is Rule)
			{
				// not valid
				return false;
			}

			if (node.Tag is Condition)
			{
				Condition condition = node.Tag as Condition;
				Condition[] conditions = node.Parent.Tag as Condition[];

				if (condition == null)
				{
					throw new Exception();
				}

				if (conditions == null)
				{
					throw new Exception();
				}

				// duplicate the condition list
				List<Condition> list = new List<Condition>();
				foreach (Condition c in conditions)
				{
					list.Add(c);
				}

				// find the index of the target condition
				int index = 0;
				for (; index < conditions.Length; ++index)
				{
					if (conditions[index].Id == condition.Id)
					{
						break;
					}
				}

				// this should never happen
				if (index >= conditions.Length)
				{
					throw new Exception();
				}

				// now, delete the entry from the list, and the node from the tree view
				list.RemoveAt(index);

				// now, set the parent's conditions
				if (node.Parent.Parent.Tag is Rule)
				{
					Rule rule = node.Parent.Parent.Tag as Rule;
					rule.Conditions = list.ToArray();
					node.Parent.Tag = rule.Conditions;
				}
				else if (node.Parent.Parent.Tag is ConditionGroup)
				{
					ConditionGroup cg = node.Parent.Parent.Tag as ConditionGroup;
					cg.Conditions = list.ToArray();
					node.Parent.Tag = cg.Conditions;
				}
				else
				{
					throw new Exception();
				}

				node.Remove();
				return true;
			}
			else if (node.Tag is Action)
			{
				Action action = node.Tag as Action;
				Rule rule = node.Parent.Parent.Tag as Rule;

				if (action == null)
				{
					throw new Exception();
				}

				if (rule == null)
				{
					throw new Exception();
				}

				// duplicate the condition list
				List<Action> list = new List<Action>();
				foreach (Action a in rule.Actions)
				{
					list.Add(a);
				}

				// find the index of the target condition
				int index = 0;
				for (; index < rule.Actions.Length; ++index)
				{
					if (rule.Actions[index].Id == action.Id)
					{
						break;
					}
				}

				// this should never happen
				if (index >= rule.Actions.Length)
				{
					throw new Exception();
				}

				// now, delete the entry from the list, and the node from the tree view
				list.RemoveAt(index);

				// now, set the parent's conditions
				rule.Actions = list.ToArray();
				node.Remove();
				return true;
			}
			else
			{
				return false;
			}
		}

		//=========================================================================================
		// drag and drop: https://learn.microsoft.com/en-us/dotnet/desktop/winforms/advanced/walkthrough-performing-a-drag-and-drop-operation-in-windows-forms
		//=========================================================================================
		private TreeNode _draggingNode = null;
		private Rule _draggingRule = null;
		private bool _draggingActive = false;

		private void _settingsTreeView_ItemDrag(object sender, ItemDragEventArgs e)
		{
			TreeView tree = sender as TreeView;

			if (tree != null)
			{
				// Move the dragged node when the left mouse button is used.  
				if (e.Button == MouseButtons.Left)
				{
					//
					// find the node we're dragging...
					//
					this._draggingNode = e.Item as TreeNode;
					TreeNode ruleNode = GetNodeRuleNode(this._draggingNode);
					if (ruleNode != null)
					{
						this._draggingRule = ruleNode.Tag as Rule;
						this._draggingActive = true;
						DoDragDrop(e.Item, DragDropEffects.Move);
					}
				}
			}
		}

		//=========================================================================================
		//=========================================================================================
		private void _settingsTreeView_DragEnter(object sender, DragEventArgs e)
		{
			//TreeView tree = sender as TreeView;
			//if (tree != null)
			//{
			//	//var d = e.Data;
			//	e.Effect = e.AllowedEffect;
			//	//e.Effect = DragDropEffects.Move;
			//}
		}

		//=========================================================================================
		//=========================================================================================
		private void _settingsTreeView_DragLeave(object sender, EventArgs e)
		{
			//TreeView tree = sender as TreeView;
			//if (tree != null)
			//{
			//}
		}

		//=========================================================================================
		//=========================================================================================
		private void _settingsTreeView_DragDrop(object sender, DragEventArgs e)
		{
			TreeView tree = sender as TreeView;
			if ((this._draggingActive) && (tree != null))
			{
				// Retrieve the client coordinates of the mouse position.
				Point targetPoint = tree.PointToClient(new Point(e.X, e.Y));

				// Select the node at the mouse position.
				TreeNode node = tree.GetNodeAt(targetPoint);

				Condition condition = _draggingNode.Tag as Condition;

				if (condition != null)
				{
					if (this.InsertConditionBeforeNode(node, condition))
					{
						this.DeleteConditionOrActionAtNode(_draggingNode);
					}
				}
			}
		}

		//=========================================================================================
		//=========================================================================================
		private void _settingsTreeView_DragOver(object sender, DragEventArgs e)
		{
			TreeView tree = sender as TreeView;
			if (tree != null)
			{
				// Retrieve the client coordinates of the mouse position.  
				Point targetPoint = tree.PointToClient(new Point(e.X, e.Y));

				// Select the node at the mouse position.  
				TreeNode node = tree.GetNodeAt(targetPoint);

				if (node == null)
				{
					e.Effect = DragDropEffects.None;
					return;
				}

				if (node.Tag == null)
				{
					e.Effect = DragDropEffects.None;
					return;
				}

				if (node.Tag is Condition)
				{
					// valid
					if (node.Tag is Rule)
					{
						// not valid
						e.Effect = DragDropEffects.None;
						return;
					}
				}
				else
				{
					// not valid
					e.Effect = DragDropEffects.None;
					return;
				}

				if (node != null)
				{
					Logger.the.WriteLineFormat("Dragging over an item with text: {0}", node.Text);
				}

				// get that node's rule
				TreeNode thisRuleNode = GetNodeRuleNode(node);

				// now, see if it's something we CAN drop on
				if ((thisRuleNode != null) && (this._draggingNode != null))
				{
					Rule thisRule = thisRuleNode.Tag as Rule;

					if ((thisRule != null) && (this._draggingRule != null))
					{
						if (thisRule.Index == this._draggingRule.Index)
						{
							tree.SelectedNode = node;
							e.Effect = DragDropEffects.Move;
						}
						else
						{
							e.Effect = DragDropEffects.None;
						}
					}
					else
					{
						e.Effect = DragDropEffects.None;
					}
				}
				else
				{
					e.Effect = DragDropEffects.None;
				}
			}
		}


		//=========================================================================================
		//=========================================================================================
		private void _settingsTreeView_QueryContinueDrag(object sender, QueryContinueDragEventArgs e)
		{
			//TreeView tree = sender as TreeView;
			//if (tree != null)
			//{
			//	e.Action = DragAction.Continue;
			//}
		}


		//=========================================================================================
		//=========================================================================================
		private void _settingsTreeView_GiveFeedback(object sender, GiveFeedbackEventArgs e)
		{
			nop();
		}


		//=========================================================================================
		//=========================================================================================
		private void FixTreeNodeRecipientName(TreeNode node)
		{
			if (node.IsExpanded)
			{
				foreach (TreeNode child in node.Nodes)
				{
					if (child.Tag != null)
					{
						if (child.Tag is RecipientCondition)
						{
							RecipientCondition c = child.Tag as RecipientCondition;
							nop(c);
						}
					}

					FixTreeNodeRecipientName(child);
				}
			}

			if (node.Tag != null)
			{
				if (node.Text == "Unknown")
				{
					if (node.Tag is RecipientCondition)
					{
						SetNodeText(node);
					}
					else if (node.Tag is SenderCondition)
					{
						SetNodeText(node);
					}
				}
			}
		}

		private void _settingsTreeView_BeforeExpand(object sender, TreeViewCancelEventArgs e)
		{
		}

		private void AddConditions(TreeNode conditionsGroupNode)
		{
			ConditionGroup conditionsGroup = conditionsGroupNode.Tag as ConditionGroup;

			if (conditionsGroup != null)
			{
				TreeNode conditionsNode = conditionsGroupNode.Nodes.Add("");
				if (conditionsGroup.Conditions == null)
				{
					conditionsGroup.Conditions = new Condition[0];
				}
				conditionsNode.Tag = conditionsGroup.Conditions;
				SetNodeText(conditionsNode);

				foreach (var condition in conditionsGroup.Conditions)
				{
					TreeNode conditionNode = conditionsNode.Nodes.Add("");
					conditionNode.Tag = condition;
					SetNodeText(conditionNode);

					if (condition is ConditionGroup)
					{
						AddConditions(conditionNode);
					}
				}
			}
		}

		private void _settingsTreeView_MouseClick(object sender, MouseEventArgs e)
		{
		}

		private ContextMenu _ruleContextMenu;
		private ContextMenu _conditionsContextMenu;
		private ContextMenu _groupConditionsContextMenu;
		private ContextMenu _actionsContextMenu;
		private ContextMenu _conditionContextMenu;
		private ContextMenu _actionContextMenu;

		private void _settingsTreeView_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				// right-clicked...
				TreeNode node = e.Node;
				//ContextMenu mnu = new ContextMenu();

				if (node.Tag is Rule)
				{
					if (this._ruleContextMenu == null)
					{
						this._ruleContextMenu = new ContextMenu();
						this._ruleContextMenu.MenuItems.Add("Edit Rule").Click += MenuItem_Click;
					}

					this._ruleContextMenu.MenuItems[0].Tag = node;

					this._ruleContextMenu.Show(this._settingsTreeView, e.Location);
					//MessageBox.Show("Recipient not found");
				}
				else if (node.Tag is Rule[])
				{
				}
				else if (node.Tag is Condition[])
				{
					if (this._conditionsContextMenu == null)
					{
						this._conditionsContextMenu = new ContextMenu();
						this._conditionsContextMenu.MenuItems.Add("New Recipient Condition").Click += MenuItem_Click;
						this._conditionsContextMenu.MenuItems.Add("New Sender Condition").Click += MenuItem_Click;
						this._conditionsContextMenu.MenuItems.Add("New Subject Condition").Click += MenuItem_Click;
						this._conditionsContextMenu.MenuItems.Add("New Body Condition").Click += MenuItem_Click;
						this._conditionsContextMenu.MenuItems.Add("New Condition Group").Click += MenuItem_Click;
					}

					this._conditionsContextMenu.MenuItems[0].Tag = node;
					this._conditionsContextMenu.MenuItems[1].Tag = node;
					this._conditionsContextMenu.MenuItems[2].Tag = node;
					this._conditionsContextMenu.MenuItems[3].Tag = node;
					this._conditionsContextMenu.MenuItems[4].Tag = node;

					this._conditionsContextMenu.Show(this._settingsTreeView, e.Location);
				}
				else if (node.Tag is Action[])
				{
					if (this._actionsContextMenu == null)
					{
						this._actionsContextMenu = new ContextMenu();
						this._actionsContextMenu.MenuItems.Add("New Move Action").Click += MenuItem_Click;
						this._actionsContextMenu.MenuItems.Add("New Mark as Read Action").Click += MenuItem_Click;
						//this._actionsContextMenu.MenuItems.Add("New Tag Category Action").Click += MenuItem_Click;
						//this._actionsContextMenu.MenuItems.Add("New Forward Action").Click += MenuItem_Click;
						//this._actionsContextMenu.MenuItems.Add("New Fix Links Action").Click += MenuItem_Click;
						//this._actionsContextMenu.MenuItems.Add("New Build Move Action").Click += MenuItem_Click;
					}

					this._actionsContextMenu.MenuItems[0].Tag = node;
					this._actionsContextMenu.MenuItems[1].Tag = node;
					this._actionsContextMenu.MenuItems[2].Tag = node;
					this._actionsContextMenu.MenuItems[3].Tag = node;
					this._actionsContextMenu.MenuItems[4].Tag = node;
					this._actionsContextMenu.MenuItems[5].Tag = node;

					this._actionsContextMenu.Show(this._settingsTreeView, e.Location);
				}
				else if (node.Tag is ConditionGroup)
				{
					if (this._groupConditionsContextMenu == null)
					{
						this._groupConditionsContextMenu = new ContextMenu();
						this._groupConditionsContextMenu.MenuItems.Add("Delete Condition").Click += MenuItem_Click;
					}

					this._groupConditionsContextMenu.MenuItems[0].Tag = node;
					this._groupConditionsContextMenu.Show(this._settingsTreeView, e.Location);
				}
				else if (node.Tag is Condition)
				{
					if (this._conditionContextMenu == null)
					{
						this._conditionContextMenu = new ContextMenu();
						this._conditionContextMenu.MenuItems.Add("Edit Condition").Click += MenuItem_Click;
						this._conditionContextMenu.MenuItems.Add("Delete Condition").Click += MenuItem_Click;
					}

					this._conditionContextMenu.MenuItems[0].Tag = node;
					this._conditionContextMenu.MenuItems[1].Tag = node;

					this._conditionContextMenu.Show(this._settingsTreeView, e.Location);
				}
				else if (node.Tag is Action)
				{
					if (this._actionContextMenu == null)
					{
						this._actionContextMenu = new ContextMenu();
						this._actionContextMenu.MenuItems.Add("Edit Action").Click += MenuItem_Click;
						this._actionContextMenu.MenuItems.Add("Delete Action").Click += MenuItem_Click;
					}

					this._actionContextMenu.MenuItems[0].Tag = node;
					this._actionContextMenu.MenuItems[1].Tag = node;

					this._actionContextMenu.Show(this._settingsTreeView, e.Location);
				}
				else
				{
					//throw new Exception("Unknown type");
				}
			}
		}

		//=========================================================================================
		// Get a node's "Rule" node (the ancestor node that is a rule)
		//=========================================================================================
		private static TreeNode GetNodeRuleNode(TreeNode node)
		{
			TreeNode tmpNode=node;

			while ((tmpNode!=null) && (!(tmpNode.Tag is Rule)))
			{
				tmpNode = tmpNode.Parent;
			}

			return tmpNode;
		}

		private void newRequestOnNode(TreeNode node, string text)
		{
			if (node.Tag is Condition[])
			{
				Condition[] conditions = node.Tag as Condition[];
				TreeNode ruleNode = node.Parent;
				Rule rule = ruleNode.Tag as Rule;

				if (ruleNode == null)
				{
					nop();
				}
				else
				{
					if (text == "New Recipient Condition")
					{
						List<Condition> list = new List<Condition>();
						foreach (Condition c in rule.Conditions)
						{
							list.Add(c);
						}

						RecipientCondition rc = new RecipientCondition();
						rc.Recipient = "Dummy";
						list.Add(rc);

						// You can convert it back to an array if you would like to
						rule.Conditions = list.ToArray();

						TreeNode newNode = node.Nodes.Add("Unknown Recipient");
						newNode.Tag = rc;
						SetNodeText(newNode);
					}
					else if (text == "New Sender Condition")
					{
						nop();
					}
					else if (text == "New Subject Condition")
					{
						nop();
					}
					else if (text == "New Body Condition")
					{
						nop();
					}
					else if (text == "New Condition Group")
					{
						//
						// the node's parent is a condition group... see what it is..
						//
						ConditionGroup parent = node.Parent.Tag as ConditionGroup;

						List<Condition> list = new List<Condition>();
						foreach (Condition c in parent.Conditions)
						{
							list.Add(c);
						}

						if (parent == null)
						{
							throw new Exception();
						}

						ConditionGroup condition = new ConditionGroup();
						condition.Conditions = new Condition[0];


						if (parent.Operator == ConditionGroup.GroupingOperator.And)
						{
							condition.Operator = ConditionGroup.GroupingOperator.Or;
						}
						else
						{
							condition.Operator = ConditionGroup.GroupingOperator.And;
						}
						list.Add(condition);

						// You can convert it back to an array if you would like to
						parent.Conditions = list.ToArray();

						TreeNode newNode = node.Nodes.Add("");
						newNode.Tag = condition;

						TreeNode newConditionsNode = newNode.Nodes.Add("");
						newConditionsNode.Tag = condition.Conditions;

						SetNodeText(newConditionsNode);
						SetNodeText(newNode);
					}
					else
					{
						throw new Exception("Unknown sender");
					}
				}
			}
			else if (node.Tag is Action[])
			{
				Action[] actions = node.Tag as Action[];

				if (text == "New Move Action")
				{
					var folder = this._application.Session.PickFolder();
					MoveAction action = new MoveAction();
					action.FolderName = folder.EntryID;
					action.FolderPath = folder.FullFolderPath;
					action.Folder = folder;

					Rule rule = node.Parent.Tag as Rule;

					if (rule == null)
					{
						throw new Exception();
					}

					List<Action> list = new List<Action>();
					foreach (Action a in rule.Actions)
					{
						list.Add(a);
					}

					list.Add(action);

					// You can convert it back to an array if you would like to
					rule.Actions = list.ToArray();

					TreeNode newNode = node.Nodes.Add("");
					newNode.Tag = action;
					SetNodeText(newNode);
				}
				else if (text == "New Mark as Read Action")
				{
					MarkAsReadAction action = new MarkAsReadAction();

					Rule rule = node.Parent.Tag as Rule;

					if (rule == null)
					{
						throw new Exception();
					}

					List<Action> list = new List<Action>();
					foreach (Action a in rule.Actions)
					{
						list.Add(a);
					}

					list.Add(action);

					// You can convert it back to an array if you would like to
					rule.Actions = list.ToArray();

					TreeNode newNode = node.Nodes.Add("");
					newNode.Tag = action;
					SetNodeText(newNode);
				}
				else if (text == "New Tag Category Action")
				{
					nop();
				}
				else if (text == "New Forward Action")
				{
					nop();
				}
				else if (text == "New Fix Links Action")
				{
					nop();
				}
				else if (text == "New Build Move Action")
				{
					nop();
				}
				else
				{
					throw new Exception("Unknown sender");
				}
			}
			else if (node.Tag is Rule[])
			{
				nop();
			}
			else if (node.Tag is Rule)
			{
				if (text == "Delete Rule")
				{
					nop();
				}
				else
				{
					nop();
				}
			}
			else if (node.Tag is Condition)
			{
				if (text == "Delete Condition")
				{
					nop();
				}
				else
				{
					nop();
				}
			}
			else if (node.Tag is Action)
			{
				if (text == "Delete Action")
				{
					nop();
				}
				else
				{
					nop();
				}
			}
		}


		private static string GetSerializedString(object o)
		{
			XmlSerializer ser = new XmlSerializer(o.GetType());
			MemoryStream stream = new MemoryStream();
			ser.Serialize(stream, o, new XmlSerializerNamespaces(new XmlQualifiedName[]{new XmlQualifiedName("")}));
			stream.Position = 0;
			return System.Text.Encoding.ASCII.GetString(stream.ToArray());
		}

		private void deleteNode(TreeNode node)
		{
			DialogResult dr = DialogResult.No;

			if (node.Tag is Condition)
			{
				dr = MessageBox.Show("Are you sure you want to delete this condition?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			}
			else if (node.Tag is Action)
			{
				dr = MessageBox.Show("Are you sure you want to delete this action?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
			}

			if (dr == DialogResult.Yes)
			{
				DeleteConditionOrActionAtNode(node);
			}
		}

		private void editNode(TreeNode node)
		{
			if (node.Tag is Rule)
			{
				Rule rule = node.Tag as Rule;

				RuleDialog dialog = new RuleDialog(rule);
				NativeWindow mainWindow = new NativeWindow();
				mainWindow.AssignHandle(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);
				DialogResult dialogResult = dialog.ShowDialog(mainWindow);
				if (dialogResult == DialogResult.OK)
				{
					// process save... find the container of this rule
					string ruleToFind = GetSerializedString(rule);
					Rule[] rules = this._settings.Rules;
					int ruleIndex = 0;

					while (ruleIndex < rules.Length)
					{
						string tmpRule = GetSerializedString(rules[ruleIndex]);

						if (tmpRule == ruleToFind)
						{
							break;
						}

						++ruleIndex;
					}

					if (ruleIndex < rules.Length)
					{
						Logger.the.WriteLineFormat("Found rule index {0}", ruleIndex);
						rules[ruleIndex] = dialog.Rule;
						node.Tag = rules[ruleIndex];
					}
					else
					{
						Logger.the.WriteLineFormat("Did not find the rule index");
					}

					SetNodeTextRecursive(node);
				}
				mainWindow.ReleaseHandle();
			}
			else if (node.Tag is Condition[])
			{
				//throw new Exception();
			}
			else if (node.Tag is Condition)
			{
				if (node.Tag is RecipientCondition)
				{
					RecipientCondition condition = node.Tag as RecipientCondition;

					ConditionDialog dialog = new ConditionDialog(this._application, condition);
					NativeWindow mainWindow = new NativeWindow();
					mainWindow.AssignHandle(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);
					DialogResult dialogResult = dialog.ShowDialog(mainWindow);
					if (dialogResult == DialogResult.OK)
					{
						//
						// get the parent to determine which "Condition" we need to fix...
						//
						Condition[] conditions = node.Parent.Tag as Condition[];
						if (conditions == null)
						{
							throw new Exception("Parent isn't a condtion array!");
						}

						//
						// get the index of this condition
						//
						int conditionIndex = 0;
						while (conditionIndex < conditions.Length)
						{
							Condition sc = conditions[conditionIndex] as Condition;

							string s1 = GetSerializedString(sc);
							string s2 = GetSerializedString(condition);

							if (s1 == s2)
							{
								break;
							}

							++conditionIndex;
						}

						if (conditionIndex < conditions.Length)
						{
							conditions[conditionIndex] = dialog.RecipientCondition;
							node.Tag = conditions[conditionIndex];
						}

						SetNodeText(node);
					}
					mainWindow.ReleaseHandle();
				}
				else if (node.Tag is SenderCondition)
				{
					SenderCondition condition = node.Tag as SenderCondition;

					ConditionDialog dialog = new ConditionDialog(this._application, condition);
					NativeWindow mainWindow = new NativeWindow();
					mainWindow.AssignHandle(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);
					DialogResult dialogResult = dialog.ShowDialog(mainWindow);
					if (dialogResult == DialogResult.OK)
					{
						//
						// get the parent to determine which "Condition" we need to fix...
						//
						Condition[] conditions = node.Parent.Tag as Condition[];

						if (conditions == null)
						{
							throw new Exception("Parent isn't a condtion array!");
						}

						int conditionIndex = 0;
						while (conditionIndex < conditions.Length)
						{
							Condition sc = conditions[conditionIndex] as Condition;

							string s1 = GetSerializedString(sc);
							string s2 = GetSerializedString(condition);

							if (s1 == s2)
							{
								break;
							}

							++conditionIndex;
						}

						if (conditionIndex < conditions.Length)
						{
							conditions[conditionIndex] = dialog.SenderCondition;
							node.Tag = conditions[conditionIndex];
						}

						SetNodeText(node);
					}
					mainWindow.ReleaseHandle();
				}
				else if (node.Tag is SubjectCondition)
				{










					SubjectCondition condition = node.Tag as SubjectCondition;

					ConditionDialog dialog = new ConditionDialog(this._application, condition);
					NativeWindow mainWindow = new NativeWindow();
					mainWindow.AssignHandle(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);
					DialogResult dialogResult = dialog.ShowDialog(mainWindow);
					if (dialogResult == DialogResult.OK)
					{
						//
						// get the parent to determine which "Condition" we need to fix...
						//
						Condition[] conditions = node.Parent.Tag as Condition[];

						if (conditions == null)
						{
							throw new Exception("Parent isn't a condtion array!");
						}

						int conditionIndex = 0;
						while (conditionIndex < conditions.Length)
						{
							Condition sc = conditions[conditionIndex] as Condition;

							string s1 = GetSerializedString(sc);
							string s2 = GetSerializedString(condition);

							if (s1 == s2)
							{
								break;
							}

							++conditionIndex;
						}

						if (conditionIndex < conditions.Length)
						{
							conditions[conditionIndex] = dialog.SubjectCondition;
							node.Tag = conditions[conditionIndex];
						}

						SetNodeText(node);
					}
					mainWindow.ReleaseHandle();





















				}
				else if (node.Tag is BodyCondition)
				{


					BodyCondition condition = node.Tag as BodyCondition;

					ConditionDialog dialog = new ConditionDialog(this._application, condition);
					NativeWindow mainWindow = new NativeWindow();
					mainWindow.AssignHandle(System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle);
					DialogResult dialogResult = dialog.ShowDialog(mainWindow);
					if (dialogResult == DialogResult.OK)
					{
						//
						// get the parent to determine which "Condition" we need to fix...
						//
						Condition[] conditions = node.Parent.Tag as Condition[];

						if (conditions == null)
						{
							throw new Exception("Parent isn't a condtion array!");
						}

						int conditionIndex = 0;
						while (conditionIndex < conditions.Length)
						{
							Condition sc = conditions[conditionIndex] as Condition;

							string s1 = GetSerializedString(sc);
							string s2 = GetSerializedString(condition);

							if (s1 == s2)
							{
								break;
							}

							++conditionIndex;
						}

						if (conditionIndex < conditions.Length)
						{
							conditions[conditionIndex] = dialog.BodyCondition;
							node.Tag = conditions[conditionIndex];
						}

						SetNodeText(node);

					}
					mainWindow.ReleaseHandle();


				}
				else if (node.Tag is ConditionGroup)
				{
					// do nothing
				}
				else
				{
				}


			}
			else if (node.Tag is Action[])
			{
				//throw new Exception();
			}
			else if (node.Tag is Action)
			{
				if (node.Tag is MoveAction)
				{
					MoveAction moveAction = node.Tag as MoveAction;
					var folder = this._application.Session.PickFolder();
					if (folder != null)
					{
						moveAction.Folder = folder;
						moveAction.FolderName = folder.EntryID;
						moveAction.FolderPath = folder.FullFolderPath;
						SetNodeText(node);
					}
					//this.nop(folder);
				}
				nop();
			}
			else
			{
				nop();
			}
		}

		private void _settingsTreeView_MouseDoubleClick(object sender, MouseEventArgs e)
		{
			if (sender is MenuItem)
			{
				nop();
			}
			else if (sender is TreeNode)
			{
				nop();
			}
			else if (sender is TreeView)
			{
				nop();
			}
			else
			{
				nop();
			}
		}

		private void _settingsTreeView_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
		{
			if (sender is MenuItem)
			{
				nop();
			}
			else if (sender is TreeNode)
			{
				nop();
			}
			else if (sender is TreeView)
			{
				TreeView treeView = sender as TreeView;
				editNode(treeView.SelectedNode);
			}
			else
			{
				nop();
			}
		}

		private void MenuItem_Click(object sender, EventArgs e)
		{
			if (sender is MenuItem)
			{
				MenuItem m = sender as MenuItem;
				TreeNode node = m.Tag as TreeNode;

				string abbrevText = m.Text.ToLower().Trim();

				if ((abbrevText.Length >= 3) && (abbrevText.Substring(0,3) == "new"))
				{
					newRequestOnNode(node, m.Text);
				}
				else if (((abbrevText.Length >= 4) && (abbrevText.Substring(0,4) == "edit")))
				{
					editNode(node);
				}
				else if (((abbrevText.Length >= 6) && (abbrevText.Substring(0,6) == "delete")))
				{
					deleteNode(node);
				}
				else
				{
					throw new Exception("Unknown menu option");
				}
			}
		}

		private void SettingsDialog_Load(object sender, EventArgs e)
		{
			this._settingsTreeView.Nodes[0].Expand();
			TreeNode node = this._settingsTreeView.SelectedNode;
			if (null != node)
			{
				node.ExpandAll();
				this._settingsTreeView.Select();
			}
		}

		private void _settingsTreeView_AfterExpand(object sender, TreeViewEventArgs e)
		{
			TreeView root = sender as TreeView;

			if (root != null)
			{
				foreach (TreeNode node in root.Nodes)
				{
					FixTreeNodeRecipientName(node);
				}
			}
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
