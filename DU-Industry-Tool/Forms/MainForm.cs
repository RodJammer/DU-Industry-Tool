﻿// ReSharper disable LocalizableElement
// ReSharper disable RedundantUsingDirective
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Globalization;
using System.Text;
using System.Windows.Forms;
using ClosedXML.Excel;
using Krypton.Toolkit;
using Krypton.Navigator;
using Krypton.Workspace;

namespace DU_Industry_Tool
{
    public partial class MainForm : KryptonForm
    {
        private readonly IndustryManager _manager;
        private readonly MarketManager _market;
        private bool _marketFiltered;
        private TextBox _costDetailsLabel;
        private readonly List<string> _breadcrumbs = new List<string>();
        private FlowLayoutPanel _costDetailsPanel;
        private FlowLayoutPanel _infoPanel;
        private bool _navUpdating;
        private int _overrideQty;

        public MainForm(IndustryManager manager)
        {
            InitializeComponent();

            CultureInfo.CurrentCulture = new CultureInfo("en-us");
            QuantityBox.SelectedIndex = 0;

            // Setup the trees. One recipe on each main node
            _manager = manager;

            _market = new MarketManager();
            kryptonPage1.Flags = 0;
            kryptonPage1.ClearFlags(KryptonPageFlags.DockingAllowDocked);
            kryptonPage1.ClearFlags(KryptonPageFlags.DockingAllowClose);
            var sortedRecipes = new List<string>();
            treeView.AfterSelect += Treeview_AfterSelect;
            treeView.NodeMouseClick += Treeview_NodeClick;
            treeView.BeginUpdate();
            foreach(var group in manager.Groupnames)
            {
                var groupNode = new TreeNode(group);
                foreach(var recipe in manager.Recipes.Where(x => x.Value?.ParentGroupName?.Equals(group, StringComparison.CurrentCultureIgnoreCase) == true)
                            .OrderBy(r => r.Value.Level).ThenBy(r => r.Value.Name)
                            .Select(x => x.Value))
                {
                    sortedRecipes.Add(recipe.Name);
                    var recipeNode = new TreeNode(recipe.Name)
                    {
                        Tag = recipe
                    };
                    recipe.Node = recipeNode;
                    groupNode.Nodes.Add(recipeNode);
                }
                treeView.Nodes.Add(groupNode);
            }
            treeView.EndUpdate();
            sortedRecipes.Sort();
            sortedRecipes.ForEach(x => SearchBox.Items.Add(x));
            OnMainformResize(null, null);
        }

        private void Treeview_NodeClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Node != null && treeView.SelectedNode == e.Node)
            {
                SelectRecipe(e.Node);
            }
        }

        private void Treeview_AfterSelect(object sender, TreeViewEventArgs e)
        {
            SelectRecipe(e?.Node);
        }

        private void SelectRecipe(TreeNode e)
        {
            if (!(e?.Tag is SchematicRecipe recipe))
            {
                return;
            }

            if (_breadcrumbs.Count == 0 || _breadcrumbs.LastOrDefault() != recipe.Name)
            {
                _breadcrumbs.Add(recipe.Name);
            }
            PreviousButton.Enabled = _breadcrumbs.Count > 0;

            // Display recipe info for the thing they have selected
            SearchBox.Text = recipe.Name;
            _navUpdating = true;
            ContentDocument newDoc = null;
            try
            {
                newDoc = NewDocument(recipe.Name);
                if (newDoc == null || _infoPanel == null) return;
            }
            finally
            {
                _navUpdating = false;
            }
            _infoPanel.Controls.Clear();

            var lbl = AddLabel(_infoPanel.Controls, recipe.Name + (_manager.ProductionListMode ? "" : $" (T{recipe.Level})"), FontStyle.Bold);
            lbl.Anchor = (AnchorStyles)(AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right);
            lbl.Font = new Font(_infoPanel.Font.FontFamily, 12f, FontStyle.Bold);
            lbl.Padding = new Padding(2, 5, 4, 2);
            lbl.Height = 30;
            lbl.Width = 370;

            var tmp = recipe.UnitMass > 0 ? $"mass: {recipe.UnitMass:N2} " : "";
            tmp += recipe.UnitVolume > 0 ? $"volume: {recipe.UnitMass:N2} " : "";
            tmp += recipe.Nanocraftable ? "*nanocraftable*" : "";
            if (tmp != "")
            {
                lbl = AddLabel(_infoPanel.Controls, "Unit " + tmp);
                lbl.Padding = new Padding(4, 0, 4, 5);
            }

            _manager.CostResults = new StringBuilder();
            var cnt = 1;
            if (_manager.ProductionListMode)
            {
                _manager.ProductQuantity = 1;
            }
            else
            {
                _manager.ProductQuantity = _overrideQty > 0 ? _overrideQty : int.Parse(QuantityBox.Text);
                if (!int.TryParse(QuantityBox.Text, out cnt)) cnt = 1;
            }

            // ***** primary calculation *****
            var costToMake = _manager.GetTotalCost(recipe.Key, cnt, silent: true);
            AddFlowLabel(_infoPanel.Controls, "Cost for 1: " + costToMake.ToString("N02") + "q");

            if (!_manager.ProductionListMode)
            {
                var cost = _manager.GetBaseCost(recipe.Key);
                AddFlowLabel(_infoPanel.Controls, $"Untalented (no schematics): {cost:N1}q");
            }

            if (recipe.Time > 0)
            {
                var sp = TimeSpan.FromSeconds(recipe.Time);
                lbl = AddLabel(_infoPanel.Controls, "Time to craft:  " +
                                                    (sp.Days > 0 ? $"{sp.Days}d : " : "") +
                                                    (sp.Hours > 0 ? $"{sp.Hours}h : " : "") +
                                                    $"{sp.Minutes}m " +
                                                    (sp.Seconds > 0 ? $" : {sp.Seconds}s" : ""));
                lbl.Padding = new Padding(4, 0, 4, 5);

                var amt = (86400 / recipe.Time).ToString();
                var pnl = AddFlowLabel(_infoPanel.Controls, "", flow: FlowDirection.LeftToRight);
                AddLabel(pnl.Controls, "Per industry ");
                AddLinkedLabel(pnl.Controls, amt, recipe.Key+"#"+amt).Click += Label_Click;
                AddLabel(pnl.Controls, " / day");
                pnl.Padding = new Padding(0, 0, 0, 10);
            }

            if (!_manager.ProductionListMode)
            {
                // IDK why sometimes prices are listed as 0
                var orders = _market.MarketOrders.Values.Where(o => o.ItemType == recipe.NqId &&
                                                                    o.BuyQuantity < 0 &&
                                                                    DateTime.Now < o.ExpirationDate &&
                                                                    o.Price > 0);
                var mostRecentOrder = orders.OrderBy(o => o.Price).FirstOrDefault();
                if (mostRecentOrder?.Price > 0.00d)
                {
                    var costPanel = new FlowLayoutPanel
                    {
                        AutoSize = true,
                        FlowDirection = FlowDirection.TopDown,
                        Padding = new Padding(0)
                    };
                    AddLabel(costPanel.Controls, "Market " + mostRecentOrder.Price.ToString("N02") + "q");
                    AddLabel(costPanel.Controls, "Until " + mostRecentOrder.ExpirationDate);
                    AddLabel(costPanel.Controls, "Profit Margin ");
                    var cost = ((mostRecentOrder.Price-costToMake) / mostRecentOrder.Price);
                    AddLabel(costPanel.Controls, cost.ToString("0%"));
                    cost = (mostRecentOrder.Price - costToMake)*(86400/recipe.Time);
                    AddLabel(costPanel.Controls, "Profit/Day/Industry " + cost.ToString("N02") + "q");
                    _infoPanel.Controls.Add(costPanel);
                }
            }

            // ----- Ingredients -----
            if (recipe.Ingredients.Count > 0)
            {
                var newPanel = AddFlowLabel(_infoPanel.Controls, "Ingredients", FontStyle.Bold);
                newPanel.SuspendLayout();
                var ingGrid = new TableLayoutPanel
                {
                    ColumnCount = 2,
                    RowCount = recipe.Ingredients.Count,
                    AutoSize = true,
                    Margin = new Padding(0),
                    Padding = new Padding(0, 0, 0, 10),
                    Width = 300
                };
                ingGrid.LayoutSettings.ColumnStyles.Add(new ColumnStyle { Width = 240, SizeType = SizeType.Absolute });
                ingGrid.SuspendLayout();
                foreach (var ingredient in recipe.Ingredients)
                {
                    //var qty = ((_overrideQty == 0 ? 1 : _overrideQty) * ingredient.Quantity).ToString("0");
                    var qty = ingredient.Quantity.ToString("0");
                    AddLinkedLabel(ingGrid.Controls, ingredient.Name, ingredient.Type+"#"+qty).Click += Label_Click;
                    AddLabel(ingGrid.Controls, qty);
                }

                ingGrid.ResumeLayout(false);
                newPanel.Controls.Add(ingGrid);
                newPanel.ResumeLayout(false);
            }

            // ----- Products -----
            if (recipe.Products.Count > 0)
            {
                var newPanel = AddFlowLabel(_infoPanel.Controls, "Products", FontStyle.Bold);
                newPanel.SuspendLayout();
                newPanel.Height = Math.Min(recipe.Products.Count+2, 10) * Math.Abs(_infoPanel.Font.Height);
                var grid = new TableLayoutPanel
                {
                    ColumnCount = 2,
                    RowCount = recipe.Products.Count,
                    AutoSize = true,
                    Padding = new Padding(0, 4, 0, 10),
                    Width = 300
                };
                grid.LayoutSettings.ColumnStyles.Add(new ColumnStyle { Width = 240, SizeType = SizeType.Absolute });
                grid.SuspendLayout();
                foreach (var prod in recipe.Products)
                {
                    if (prod.Type == recipe.Key)
                        AddLabel(grid.Controls, prod.Name);
                    else
                        AddLinkedLabel(grid.Controls, prod.Name, prod.Type).Click += Label_Click;
                    //var qty = (_overrideQty == 0 ? prod.Quantity : _overrideQty).ToString("0");
                    var qty = prod.Quantity.ToString("0");
                    AddLabel(grid.Controls, qty);
                }

                grid.ResumeLayout(false);
                newPanel.Controls.Add(grid);
                newPanel.ResumeLayout(false);
            }
            _overrideQty = 0; // must reset here!

            // ----- Industry -----
            if (!string.IsNullOrEmpty(recipe.Industry))
            {
                var newPanel = AddFlowLabel(_infoPanel.Controls, "Industry", FontStyle.Bold);
                AddLinkedLabel(newPanel.Controls, recipe.Industry, recipe.Industry).Click += LabelIndustry_Click;
                lbl = AddLabel(newPanel.Controls, "*not 100% correct!",FontStyle.Italic);
                lbl.Font = new Font("Segoe UI", 8F, FontStyle.Regular, GraphicsUnit.Point, 0);
            }

            // ----- Reverse recipes -----
            if (recipe.ParentGroupName != null &&
                (recipe.ParentGroupName.EndsWith("Ore", StringComparison.InvariantCultureIgnoreCase) ||
                recipe.ParentGroupName.EndsWith("Parts", StringComparison.InvariantCultureIgnoreCase) ||
                recipe.ParentGroupName.EndsWith("Product", StringComparison.InvariantCultureIgnoreCase) ||
                recipe.ParentGroupName.EndsWith("Pure", StringComparison.InvariantCultureIgnoreCase) ||
                recipe.Name?.StartsWith("Relic Plasma", StringComparison.InvariantCultureIgnoreCase) == true))
            {
                var containedIn = _manager.Recipes.Values.Where(x =>
                    true == x.Ingredients?.Any(y => y.Name.Equals(recipe.Name, StringComparison.InvariantCultureIgnoreCase))).ToList();
                if (containedIn?.Any() == true)
                {
                    var newPanel = AddFlowLabel(_infoPanel.Controls, "Part of recipes:", FontStyle.Bold);
                    newPanel.Padding = new Padding(0, 10, 0, 0);

                    var grid = new TableLayoutPanel
                    {
                        AutoScroll = true,
                        ColumnCount = 1,
                        Padding = new Padding(0),
                        RowCount = containedIn.Count(),
                        Height = Math.Min(containedIn.Count()+1, 10) * Math.Abs(_infoPanel.Font.Height+1),
                        Width = 300,
                        VerticalScroll = { Visible = true }
                    };
                    newPanel.SuspendLayout();
                    grid.SuspendLayout();
                    foreach (var entry in containedIn)
                    {
                        AddLinkedLabel(grid.Controls, entry.Name, entry.Key).Click += Label_Click;
                    }
                    grid.ResumeLayout(false);
                    newPanel.Controls.Add(grid);
                    newPanel.ResumeLayout(false);
                }
            }

            _costDetailsPanel = new FlowLayoutPanel();
            _costDetailsPanel.SuspendLayout();
            try
            {
                _costDetailsPanel.FlowDirection = FlowDirection.TopDown;
                _costDetailsPanel.BorderStyle = BorderStyle.None;
                _costDetailsPanel.Dock = DockStyle.None;
                _costDetailsPanel.AutoSize = true;
                _costDetailsPanel.AutoScroll = true;
                _costDetailsPanel.Font = new Font("Lucida Console", 10F, FontStyle.Regular, GraphicsUnit.Point, 0);
                _costDetailsPanel.Size = new Size(600, 500);
                _costDetailsLabel = new TextBox
                {
                    AutoSize = true,
                    BorderStyle = BorderStyle.None,
                    ScrollBars = ScrollBars.Both,
                    Size = new Size(600, 500),
                    Text = _manager.CostResults.ToString(),
                    Multiline = true,
                    WordWrap = false,
                    ReadOnly = true
                };
                _costDetailsPanel.Controls.Add(_costDetailsLabel);
                _infoPanel.Controls.Add(_costDetailsPanel);
            }
            finally
            {
                _costDetailsPanel.ResumeLayout(false);
                newDoc.CostDetailsPanel = _costDetailsPanel;
            }

            _infoPanel.AutoScroll = false;
            OnMainformResize(null, null);
        }

        private FlowLayoutPanel AddFlowLabel(System.Windows.Forms.Control.ControlCollection cc,
            string lblText, FontStyle fstyle = FontStyle.Regular,
            FlowDirection flow = FlowDirection.TopDown)
        {
            var panel = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = flow,
                Padding = new Padding(0)
            };
            if (!string.IsNullOrEmpty(lblText))
            {
                var lbl = new Label
                {
                    AutoSize = true,
                    Font = new Font(_infoPanel.Font, fstyle),
                    Padding = new Padding(0),
                    Text = lblText
                };
                panel.Controls.Add(lbl);
            }
            cc.Add(panel);
            return panel;
        }

        private Label AddLabel(System.Windows.Forms.Control.ControlCollection cc, string lblText, FontStyle fstyle = FontStyle.Regular)
        {
            var lbl = new Label
            {
                AutoSize = true,
                Font = new Font(_infoPanel.Font, fstyle),
                Padding = new Padding(0),
                Text = lblText
            };
            cc.Add(lbl);
            return lbl;
        }

        private Label AddLinkedLabel(System.Windows.Forms.Control.ControlCollection cc, string lblText, string lblKey)
        {
            var lbl = AddLabel(cc, lblText, FontStyle.Underline);
            lbl.ForeColor = Color.CornflowerBlue;
            lbl.Text = lblText;
            lbl.Tag = lblKey;
            return lbl;
        }

        private void Label_Click(object sender, EventArgs e)
        {
            if (!(sender is Label label)) return;
            if (!(label.Tag is string tag)) return;
            _overrideQty = 0;
            var hashpos = tag.IndexOf('#');
            if (hashpos > 0)
            {
                var amount = tag.Substring(hashpos+1);
                if (!int.TryParse(amount, out _overrideQty))
                {
                    _overrideQty = 0;
                }
                tag = tag.Substring(0, hashpos);
            }
            if (!_manager.Recipes.ContainsKey(tag))
            {
                return;
            }
            var recipe = _manager.Recipes[tag];

            if (_breadcrumbs.Count == 0 || _breadcrumbs.LastOrDefault() != recipe.Name)
            {
                _breadcrumbs.Add(recipe.Name);
            }
            PreviousButton.Enabled = _breadcrumbs.Count > 0;

            var outerNodes = treeView.Nodes.OfType<TreeNode>();
            TreeNode targetNode = null;
            foreach(var outerNode in outerNodes)
            {
                foreach(var innerNode in outerNode.Nodes.OfType<TreeNode>())
                {
                    if (!innerNode.Text.Equals(recipe.Name, StringComparison.CurrentCultureIgnoreCase)) continue;
                    targetNode = innerNode;
                    break;
                }
                if (targetNode != null) break;
            }
            if (targetNode == null) return;
            if (treeView.SelectedNode == targetNode)
            {
                treeView.SelectedNode = null;
            }
            treeView.SelectedNode = targetNode;
            treeView.SelectedNode.EnsureVisible();
            SearchBox.Text = targetNode.Text;
        }

        private static int RecipeNameComparer(SchematicRecipe x, SchematicRecipe y)
        {
            return string.Compare(x.Name, y.Name, StringComparison.InvariantCultureIgnoreCase);
        }

        private void LabelIndustry_Click(object sender, EventArgs e)
        {
            if (!(sender is Label label)) return;
            if (!(label.Tag is string tag)) return;
            var products = _manager.Recipes.Where(x => x.Value.Industry.Equals(tag, StringComparison.InvariantCultureIgnoreCase))
                                .Select(x => x.Value).ToList();
            if (products?.Any() != true)
            {
                return;
            }

            products.Sort(RecipeNameComparer);

            const string title = "Industry Products";

            var page = kryptonNavigator1.Pages.FirstOrDefault(x => x.Text.Equals(title, StringComparison.InvariantCultureIgnoreCase));
            if (page == null)
            {
                page = NewPage(title, null);
                kryptonNavigator1.Pages.Insert(0, page);
            }
            if (page == null) return;
            kryptonNavigator1.SelectedPage = page;
            page.Controls.Clear();

            var panel = AddFlowLabel(page.Controls, tag+" produces:", FontStyle.Bold);
            panel.Dock = DockStyle.Top;

            var grid = new TableLayoutPanel
            {
                AutoScroll = true,
                AutoSize = true,
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = products.Count,
                Margin = new Padding(0, 22, 0, 0),
                Padding = new Padding(0, 22, 0, 10)
            };
            page.Controls.Add(grid);

            grid.SuspendLayout();
            foreach (var prod in products)
            {
                AddLinkedLabel(grid.Controls, prod.Name, prod.Key).Click += Label_Click;
            }
            grid.ResumeLayout();
        }

        private void InputOreValuesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var form = new OreValueForm(_manager);
            form.ShowDialog(this);
        }

        private void SkillLevelsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var form = new SkillForm(_manager);
            form.ShowDialog(this);
        }

        private void SchematicValuesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var form = new SchematicValueForm(_manager);
            form.ShowDialog(this);
        }

        private void PreviousButton_Click(object sender, EventArgs e)
        {
            try
            {
                if (_breadcrumbs.Count == 0) return;
                var entry = _breadcrumbs.LastOrDefault();
                if (entry == SearchBox.Text)
                {
                    _breadcrumbs.Remove(entry);
                    entry = _breadcrumbs.LastOrDefault();
                }
                if (string.IsNullOrEmpty(entry)) return;
                SearchBox.Text = entry;
                _breadcrumbs.Remove(entry);
                SearchButton_Click(sender, null);
            }
            finally
            {
                PreviousButton.Enabled = _breadcrumbs.Count > 0;
            }
        }

        private void SearchButton_Click(object sender, EventArgs e)
        {
            var searchValue = SearchBox.Text;
            if (string.IsNullOrWhiteSpace(searchValue))
            {
                PreviousButton.Enabled = _breadcrumbs.Count > 0;
                return; // Do nothing
            }

            var outerNodes = treeView.Nodes.OfType<TreeNode>();
            TreeNode firstResult = null;
            treeView.BeginUpdate();
            try
            {
                treeView.CollapseAll();
                foreach (var outerNode in outerNodes)
                {
                    foreach (var innerNode in outerNode.Nodes.OfType<TreeNode>())
                    {
                        if (innerNode.Text.IndexOf(searchValue, StringComparison.InvariantCultureIgnoreCase) < 0)
                            continue;
                        innerNode.EnsureVisible();
                        var isExact = innerNode.Text.Equals(searchValue, StringComparison.InvariantCultureIgnoreCase);
                        if (firstResult == null || isExact)
                        {
                            firstResult = innerNode;
                            if (isExact) break;
                        }
                    }
                }
                if (firstResult != null)
                {
                    treeView.SelectedNode = firstResult;
                    treeView.SelectedNode.EnsureVisible();
                }
            }
            finally
            {
                treeView.EndUpdate();
            }
            treeView.Focus();
        }

        private void QuantityBoxOnSelectionChangeCommitted(object sender, EventArgs e)
        {
            if (treeView.SelectedNode == null)
            {
                SearchButton_Click(sender, e);
            }
            else
            {
                SelectRecipe(treeView.SelectedNode);
            }
        }

        private void UpdateMarketValuesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var loadForm = new LoadingForm(_market);
            loadForm.ShowDialog(this);
            if (loadForm.DiscardOres)
            {
                // Get rid of them
                List<ulong> toRemove = new List<ulong>();
                foreach(var order in _market.MarketOrders)
                {
                    var recipe = _manager.Recipes.Values.Where(r => r.NqId == order.Value.ItemType).FirstOrDefault();
                    if (recipe != null && recipe.ParentGroupName == "Ore")
                        toRemove.Add(order.Key);

                }
                foreach (var key in toRemove)
                    _market.MarketOrders.Remove(key);
                _market.SaveData();
            }
            else
            {
                // Process them and leave them so they show in exports
                foreach (var order in _market.MarketOrders)
                {
                    var recipe = _manager.Recipes.Values.Where(r => r.NqId == order.Value.ItemType).FirstOrDefault();
                    if (recipe != null && recipe.ParentGroupName == "Ore")
                    {
                        var ore = _manager.Ores.Where(o => o.Key.ToLower() == recipe.Key.ToLower()).FirstOrDefault();
                        if (ore != null)
                        {
                            var orders = _market.MarketOrders.Values.Where(o => o.ItemType == recipe.NqId && o.BuyQuantity < 0 && DateTime.Now < o.ExpirationDate && o.Price > 0);

                            var bestOrder = orders.OrderBy(o => o.Price).FirstOrDefault();
                            if (bestOrder != null)
                                ore.Value = bestOrder.Price;
                        }
                    }

                }
                _manager.SaveOreValues();
            }
            loadForm.Dispose();
        }

        private void FilterToMarketToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_marketFiltered)
            {
                _marketFiltered = false;
                if (sender is ToolStripMenuItem tsItem) tsItem.Text = "Filter to Market";
                else
                if (sender is KryptonContextMenuItem kBtn) kBtn.Text = "Filter to Market";
                treeView.Nodes.Clear();
                foreach (var group in _manager.Recipes.Values.GroupBy(r => r.ParentGroupName))
                {
                    var groupNode = new TreeNode(group.Key);
                    foreach (var recipe in group)
                    {
                        var recipeNode = new TreeNode(recipe.Name)
                        {
                            Tag = recipe
                        };
                        recipe.Node = recipeNode;

                        groupNode.Nodes.Add(recipeNode);
                    }
                    treeView.Nodes.Add(groupNode);
                }
            }
            else
            {
                _marketFiltered = true;
                if (sender is ToolStripMenuItem tsItem) tsItem.Text = "Unfilter Market";
                    else
                if (sender is KryptonContextMenuItem kBtn) kBtn.Text = "Unfilter Market";
                treeView.Nodes.Clear();
                foreach (var group in _manager.Recipes.Values.Where(r => _market.MarketOrders.Values.Any(v => v.ItemType == r.NqId)).GroupBy(r => r.ParentGroupName))
                {
                    var groupNode = new TreeNode(group.Key);
                    foreach (var recipe in group)
                    {
                        var recipeNode = new TreeNode(recipe.Name)
                        {
                            Tag = recipe
                        };
                        recipe.Node = recipeNode;

                        groupNode.Nodes.Add(recipeNode);
                    }
                    treeView.Nodes.Add(groupNode);
                }
            }
        }

        private void ExportToSpreadsheetToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // If market filtered, only exports items with market values.
            // Exports the following:
            // Name, Cost To Make, Market Cost, Time To Make, Profit Margin (with formula),
            // Profit Per Day (with formula), Units Per Day with formula
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Price Data " + DateTime.Now.ToString("yyyy-MM-dd"));

                worksheet.Cell(1, 1).Value = "Name";
                worksheet.Cell(1, 2).Value = "Cost To Make";
                worksheet.Cell(1, 3).Value = "Market Cost";
                worksheet.Cell(1, 4).Value = "Time To Make";
                worksheet.Cell(1, 5).Value = "Profit Margin";
                worksheet.Cell(1, 6).Value = "Profit Per Day";
                worksheet.Cell(1, 7).Value = "Units Per Day";

                worksheet.Row(1).CellsUsed().Style.Font.SetBold();

                int row = 2;

                var recipes = _manager.Recipes.Values.OrderBy(x => x.Name).ToList();
                if (_marketFiltered)
                {
                    recipes = _manager.Recipes.Values.Where(r =>
                        _market.MarketOrders.Values.Any(v => v.ItemType == r.NqId)).ToList();
                }

                foreach(var recipe in recipes)
                {
                    worksheet.Cell(row, 1).Value = recipe.Name;
                    var costToMake = _manager.GetTotalCost(recipe.Key, silent: true);
                    worksheet.Cell(row, 2).Value = Math.Round(costToMake,2);

                    var orders = _market.MarketOrders.Values.Where(o => o.ItemType == recipe.NqId && o.BuyQuantity < 0 && DateTime.Now < o.ExpirationDate && o.Price > 0);

                    var mostRecentOrder = orders.OrderBy(o => o.Price).FirstOrDefault();
                    var cost = mostRecentOrder?.Price ?? 0d;
                    worksheet.Cell(row, 3).Value = cost;
                    worksheet.Cell(row, 4).Value = recipe.Time;
                    worksheet.Cell(row, 5).FormulaR1C1 = "=((R[0]C[-2]-R[0]C[-3])/R[0]C[-2])";
                    //worksheet.Cell(row, 5).Value = cost = ((mostRecentOrder.Price - costToMake) / mostRecentOrder.Price);
                    //worksheet.Cell(row, 5).FormulaR1C1 = "=IF((R[0]C[-2]<>0),(R[0]C[-2]-R[0]C[-3])/R[0]C[-2],0)";
                    //cost = (mostRecentOrder.Price - costToMake)*(86400/recipe.Time);
                    worksheet.Cell(row, 6).FormulaR1C1 = "=(R[0]C[-3]-R[0]C[-4])*(86400/R[0]C[-2])";
                    worksheet.Cell(row, 7).FormulaR1C1 = "=86400/R[0]C[-3]";
                    row++;
                }
                worksheet.Range("A1:G1").Style.Font.Bold = true;
                worksheet.ColumnsUsed().AdjustToContents(1, 50);
                workbook.SaveAs("Item Export " + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx");
                MessageBox.Show("Exported to Item Export " + DateTime.Now.ToString("yyyy-MM-dd") + ".xlsx in the same folder as the exe!");
            }
        }

        private void FactoryBreakdownForSelectedToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Exports an excel sheet with info about how to setup the factory for the selected recipe (aborts if no recipe selected)
            if (!(treeView.SelectedNode?.Tag is SchematicRecipe recipe)) return;
            // Shows the amount of required components, amount per day required, amount per day per industry, and the number of industries you need of that component to provide for 1 of the parent
            // The number of parent parts can be put in as a value
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Factory");
                worksheet.Cell(1, 1).Value = "Number of industries producing " + recipe.Name;
                worksheet.Cell(1, 2).Value = "Produced/day";
                worksheet.Cell(2, 1).Value = 1;
                worksheet.Cell(2, 2).FormulaR1C1 = $"=R[0]C[-1]*(86400/{recipe.Time})";

                worksheet.Cell(1, 3).Value = "Product";
                worksheet.Cell(1, 4).Value = "Required/day";
                worksheet.Cell(1, 5).Value = "Produced/day/industry";
                worksheet.Cell(1, 6).Value = "Num industries required";
                worksheet.Cell(1, 7).Value = "Actual";

                worksheet.Row(1).Style.Font.SetBold();

                int row = 2;
                var ingredients = _manager.GetIngredientRecipes(recipe.Key).OrderByDescending(i => i.Level).GroupBy(i => i.Name);
                if (!ingredients?.Any() == true) return;
                try
                {
                    foreach(var group in ingredients)
                    {
                        var groupSum = group.Sum(g => g.Quantity);
                        worksheet.Cell(row, 3).Value = group.First().Name;
                        worksheet.Cell(row, 4).FormulaA1 = $"=B2*{groupSum}";
                        double outputMult = 1;
                        var talents = _manager.Talents.Where(t => t.InputTalent == false && t.ApplicableRecipes.Contains(group.First().Name));
                        if (talents?.Any() == true)
                            outputMult += talents.Sum(t => t.Multiplier);
                        if (group.First().ParentGroupName != "Ore")
                        {
                            worksheet.Cell(row, 5).Value = (86400 / group.First().Time) * group.First().Products.First().Quantity * outputMult;
                            worksheet.Cell(row, 6).FormulaR1C1 = "=R[0]C[-2]/R[0]C[-1]";
                            worksheet.Cell(row, 7).FormulaR1C1 = "=ROUNDUP(R[0]C[-1])";
                        }
                        row++;
                    }

                    worksheet.ColumnsUsed().AdjustToContents();
                    workbook.SaveAs($"Factory Plan {recipe.Name} {DateTime.Now:yyyy-MM-dd}.xlsx");
                    MessageBox.Show($"Exported to 'Factory Plan {recipe.Name} { DateTime.Now:yyyy-MM-dd}.xlsx' in the same folder as the exe!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Sorry, an error occured during calculation!", "ERROR", MessageBoxButtons.OK);
                    Console.WriteLine(ex);
                }
            }
        }

        private void OnMainformResize(object sender, EventArgs e)
        {
            kryptonNavigator1.Left = searchPanel.Width + 0;
            kryptonNavigator1.Top = kryptonRibbon.Height + 0;
            kryptonNavigator1.Height = kryptonWorkspaceCell1.Height - 2;
            kryptonNavigator1.Width = ClientSize.Width - searchPanel.Width - 0;
            _infoPanel = null;
            _costDetailsPanel = null;
            _costDetailsLabel = null;
            if (kryptonNavigator1.SelectedPage == null) return;
            if (kryptonNavigator1.SelectedPage.Controls.Count > 0 &&
                kryptonNavigator1.SelectedPage.Controls[0] is ContentDocument xDoc)
            {
                _infoPanel = xDoc.InfoPanel;
                _costDetailsPanel = xDoc.CostDetailsPanel;
                if (_costDetailsPanel?.Controls.Count > 0)
                {
                    _costDetailsLabel = _costDetailsPanel.Controls[0] as TextBox;
                }
            }
            if (_costDetailsPanel == null || _infoPanel == null) return;
            _costDetailsPanel.SuspendLayout();
            try
            {
                _costDetailsPanel.AutoSize = false;
                _costDetailsPanel.Height = kryptonNavigator1.SelectedPage.Height - 4;
                _costDetailsPanel.Width = kryptonNavigator1.SelectedPage.Width - _costDetailsPanel.Left;
                if (_costDetailsLabel == null) return;
                _costDetailsLabel.AutoSize = false;
                _costDetailsLabel.Width  = _costDetailsPanel.Width  - 8;
                _costDetailsLabel.Height = _costDetailsPanel.Height - 8;
            }
            finally
            {
                _costDetailsPanel.ResumeLayout(false);
            }
        }

        // Krypton related stuff

        private void MainForm_Load(object sender, EventArgs e)
        {
            // Setup docking functionality
            var w = kryptonDockingManager.ManageWorkspace(kryptonDockableWorkspace);
            if (w != null)
            {
                kryptonDockingManager.ManageControl(kryptonPage1, w);
            }
            kryptonDockingManager.ManageFloating(this);

            Properties.Settings.Default.Reload();
            switch (Properties.Settings.Default.ThemeName)
            {
                case "Office2010Black":
                    RbOffice2010Black_Click(null, null);
                    break;
                case "Office2010Silver":
                    RbOffice2010BSilver_Click(null, null);
                    break;
                default:
                    RbOffice2010Blue_Click(null, null);
                    break;
            }

            // Do not allow the left-side page to be closed or made auto hidden/docked
            kryptonPage1.ClearFlags(KryptonPageFlags.DockingAllowAutoHidden |
                            KryptonPageFlags.DockingAllowDocked |
                            KryptonPageFlags.DockingAllowClose);

            QuantityBox.SelectionChangeCommitted += QuantityBoxOnSelectionChangeCommitted;
            kryptonNavigator1.SelectedPageChanged += KryptonNavigator1OnSelectedPageChanged;
            RbnBtnProductionList.Click += RbnBtnProductionList_Click;
            OnMainformResize(sender, e);
            this.Resize += OnMainformResize;
        }

        private static KryptonPage NewPage(string name, Control content)
        {
            var p = new KryptonPage(name)
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Flags = 0
            };
            p.SetFlags(KryptonPageFlags.DockingAllowDocked | KryptonPageFlags.DockingAllowClose);
            if (content == null) return p;
            content.Dock = DockStyle.Fill;
            p.Controls.Add(content);
            return p;
        }

        private ContentDocument NewDocument(string title = null)
        {
            _infoPanel = null;
            _costDetailsPanel = null;
            if (kryptonNavigator1 == null) return null;
            var oldPage = kryptonNavigator1.Pages.FirstOrDefault(x => x.Text == title);
            if (oldPage != null)
            {
                if (oldPage.Controls.Count > 0 && oldPage.Controls[0] is ContentDocument xDoc)
                {
                    _infoPanel = xDoc.InfoPanel;
                    kryptonNavigator1.SelectedPage = oldPage;
                    return xDoc;
                }
            }
            var newDoc = new ContentDocument();
            _infoPanel = newDoc.InfoPanel;
            var page = NewPage(title ?? "Cost", newDoc);
            kryptonNavigator1.Pages.Add(page);
            kryptonNavigator1.SelectedPage = page;
            return newDoc;
        }

        private void RibbonAppButtonExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void KryptonDockableWorkspace_WorkspaceCellAdding(object sender, WorkspaceCellEventArgs e)
        {
            e.Cell.Button.CloseButtonAction = CloseButtonAction.RemovePageAndDispose;
            // Remove the context menu from the tabs bar, as it is not relevant to this sample
            e.Cell.Button.ContextButtonDisplay = ButtonDisplay.Hide;
            e.Cell.Button.NextButtonDisplay = ButtonDisplay.Hide;
            e.Cell.Button.PreviousButtonDisplay = ButtonDisplay.Hide;
        }

        private void KryptonNavigator1OnSelectedPageChanged(object sender, EventArgs e)
        {
            if(_navUpdating || !(sender is KryptonNavigator nav && nav.SelectedPage != null)) return;
            if (nav.SelectedPage.Controls.Count == 0) return;
            SearchBox.Text = nav.SelectedPage.Text;
            //SearchButton_Click(SearchButton, e);
        }

        private void RibbonButtonAboutClick(object sender, EventArgs e)
        {
            var form = new AboutForm();
            form.ShowDialog(this);
        }

        private void RbOffice2010Blue_Click(object sender, EventArgs e)
        {
            kryptonManager.GlobalPaletteMode = PaletteModeManager.Office2010Blue;
            SaveSettings();
        }

        private void RbOffice2010BSilver_Click(object sender, EventArgs e)
        {
            kryptonManager.GlobalPaletteMode = PaletteModeManager.Office2010Silver;
            SaveSettings();
        }

        private void RbOffice2010Black_Click(object sender, EventArgs e)
        {
            kryptonManager.GlobalPaletteMode = PaletteModeManager.Office2010Black;
            SaveSettings();
        }

        private void SaveSettings()
        {
            Properties.Settings.Default.ThemeName = kryptonManager.GlobalPaletteMode.ToString();
            Properties.Settings.Default.Save();
        }

        private void RbnBtnProductionList_Click(object sender, EventArgs e)
        {
            var form = new ProductionListForm(_manager);

            foreach (var entry in _manager.RecipeNames)
            {
                form.RecipeNames.Items.Add(entry);
            }
            var result = form.ShowDialog(this);
            if (result == DialogResult.Cancel) return;

            // Let Manager prepare the compound recipe which includes
            // all items' ingredients and the items as products.
            if (!_manager.PrepareProductListRecipe())
            {
                KryptonMessageBox.Show("Production list could not be prepared!", "Failure");
                return;
            }

            // TODO develop new ContentDocument/method for displaying these results?
            Calculator.Initialize();
            try
            {
                _manager.ProductionListMode = true;
                SelectRecipe(new TreeNode{
                    Text = "Production List",
                    Tag = _manager.CompoundRecipe
                });
            }
            finally
            {
                _manager.ProductionListMode = false;
            }
        }

    } // Mainform
}
