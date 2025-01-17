﻿using System;
using System.ComponentModel;
using System.Windows.Forms;
using Krypton.Navigator;

namespace DU_Industry_Tool
{
    partial class MainForm
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.kryptonRibbon = new Krypton.Ribbon.KryptonRibbon();
            this.buttonOreValues = new Krypton.Toolkit.KryptonContextMenuItem();
            this.buttonSchematicValues = new Krypton.Toolkit.KryptonContextMenuItem();
            this.buttonSkillLevels = new Krypton.Toolkit.KryptonContextMenuItem();
            this.buttonUpdateMarketValues = new Krypton.Toolkit.KryptonContextMenuItem();
            this.buttonFilterToMarket = new Krypton.Toolkit.KryptonContextMenuItem();
            this.buttonExportToSpreadsheet = new Krypton.Toolkit.KryptonContextMenuItem();
            this.buttonFactoryBreakdownForSelected = new Krypton.Toolkit.KryptonContextMenuItem();
            this.ribbonAppButtonExit = new Krypton.Toolkit.KryptonContextMenuItem();
            this.rbnTabTools = new Krypton.Ribbon.KryptonRibbonTab();
            this.kryptonRibbonGroup1 = new Krypton.Ribbon.KryptonRibbonGroup();
            this.kryptonRibbonGroupTriple1 = new Krypton.Ribbon.KryptonRibbonGroupTriple();
            this.RbnBtnProductionList = new Krypton.Ribbon.KryptonRibbonGroupButton();
            this.rbnTabThemes = new Krypton.Ribbon.KryptonRibbonTab();
            this.rbnGrpThemes = new Krypton.Ribbon.KryptonRibbonGroup();
            this.rbnTrippleThemes = new Krypton.Ribbon.KryptonRibbonGroupTriple();
            this.rbnBtnThemeBlue = new Krypton.Ribbon.KryptonRibbonGroupButton();
            this.rbnBtnThemeSilver = new Krypton.Ribbon.KryptonRibbonGroupButton();
            this.rbnBtnThemeBlack = new Krypton.Ribbon.KryptonRibbonGroupButton();
            this.kryptonRibbonTab2 = new Krypton.Ribbon.KryptonRibbonTab();
            this.kryptonRibbonGroup5 = new Krypton.Ribbon.KryptonRibbonGroup();
            this.kryptonRibbonGroupLines1 = new Krypton.Ribbon.KryptonRibbonGroupLines();
            this.kryptonRibbonGroupButton19 = new Krypton.Ribbon.KryptonRibbonGroupButton();
            this.cmdThemeBlue = new Krypton.Toolkit.KryptonCommand();
            this.kryptonDockableWorkspace = new Krypton.Docking.KryptonDockableWorkspace();
            this.kryptonPage1 = new Krypton.Navigator.KryptonPage();
            this.searchPanel = new Krypton.Toolkit.KryptonPanel();
            this.SearchBox = new Krypton.Toolkit.KryptonComboBox();
            this.SearchButton = new Krypton.Toolkit.KryptonButton();
            this.QuantityBox = new Krypton.Toolkit.KryptonComboBox();
            this.PreviousButton = new Krypton.Toolkit.KryptonButton();
            this.treeView = new System.Windows.Forms.TreeView();
            this.kryptonWorkspaceCell1 = new Krypton.Workspace.KryptonWorkspaceCell();
            this.kryptonNavigator1 = new Krypton.Navigator.KryptonNavigator();
            this.imageListSmall = new System.Windows.Forms.ImageList(this.components);
            this.kryptonDockingManager = new Krypton.Docking.KryptonDockingManager();
            this.kryptonManager = new Krypton.Toolkit.KryptonManager(this.components);
            this.kryptonRibbonGroupGallery1 = new Krypton.Ribbon.KryptonRibbonGroupGallery();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonRibbon)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonDockableWorkspace)).BeginInit();
            this.kryptonDockableWorkspace.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPage1)).BeginInit();
            this.kryptonPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.searchPanel)).BeginInit();
            this.searchPanel.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.SearchBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonWorkspaceCell1)).BeginInit();
            this.kryptonWorkspaceCell1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonNavigator1)).BeginInit();
            this.kryptonNavigator1.SuspendLayout();
            this.SuspendLayout();
            // 
            // kryptonRibbon
            // 
            this.kryptonRibbon.AllowMinimizedChange = false;
            this.kryptonRibbon.InDesignHelperMode = true;
            this.kryptonRibbon.MinimizedMode = true;
            this.kryptonRibbon.Name = "kryptonRibbon";
            this.kryptonRibbon.RibbonAppButton.AppButtonMenuItems.AddRange(new Krypton.Toolkit.KryptonContextMenuItemBase[] {
            this.buttonOreValues,
            this.buttonSchematicValues,
            this.buttonSkillLevels,
            this.buttonUpdateMarketValues,
            this.buttonFilterToMarket,
            this.buttonExportToSpreadsheet,
            this.buttonFactoryBreakdownForSelected,
            this.ribbonAppButtonExit});
            this.kryptonRibbon.RibbonAppButton.AppButtonShowRecentDocs = false;
            this.kryptonRibbon.RibbonTabs.AddRange(new Krypton.Ribbon.KryptonRibbonTab[] {
            this.rbnTabTools,
            this.rbnTabThemes,
            this.kryptonRibbonTab2});
            this.kryptonRibbon.SelectedTab = this.rbnTabTools;
            this.kryptonRibbon.ShowMinimizeButton = false;
            this.kryptonRibbon.Size = new System.Drawing.Size(1850, 136);
            this.kryptonRibbon.TabIndex = 0;
            // 
            // buttonOreValues
            // 
            this.buttonOreValues.Text = "&Ore/Plasma Values";
            this.buttonOreValues.Click += new System.EventHandler(this.InputOreValuesToolStripMenuItem_Click);
            // 
            // buttonSchematicValues
            // 
            this.buttonSchematicValues.Text = "&Schematics";
            this.buttonSchematicValues.Click += new System.EventHandler(this.SchematicValuesToolStripMenuItem_Click);
            // 
            // buttonSkillLevels
            // 
            this.buttonSkillLevels.Text = "&Talents";
            this.buttonSkillLevels.Click += new System.EventHandler(this.SkillLevelsToolStripMenuItem_Click);
            // 
            // buttonUpdateMarketValues
            // 
            this.buttonUpdateMarketValues.Text = "&Update Market Values";
            this.buttonUpdateMarketValues.Click += new System.EventHandler(this.UpdateMarketValuesToolStripMenuItem_Click);
            // 
            // buttonFilterToMarket
            // 
            this.buttonFilterToMarket.Text = "&Filter to Market";
            this.buttonFilterToMarket.Click += new System.EventHandler(this.FilterToMarketToolStripMenuItem_Click);
            // 
            // buttonExportToSpreadsheet
            // 
            this.buttonExportToSpreadsheet.Text = "&Export to CSV";
            this.buttonExportToSpreadsheet.Click += new System.EventHandler(this.ExportToSpreadsheetToolStripMenuItem_Click);
            // 
            // buttonFactoryBreakdownForSelected
            // 
            this.buttonFactoryBreakdownForSelected.Text = "&Factory Breakdown for Selected";
            this.buttonFactoryBreakdownForSelected.Click += new System.EventHandler(this.FactoryBreakdownForSelectedToolStripMenuItem_Click);
            // 
            // ribbonAppButtonExit
            // 
            this.ribbonAppButtonExit.Text = "E&xit";
            this.ribbonAppButtonExit.Click += new System.EventHandler(this.RibbonAppButtonExit_Click);
            // 
            // rbnTabTools
            // 
            this.rbnTabTools.Groups.AddRange(new Krypton.Ribbon.KryptonRibbonGroup[] {
            this.kryptonRibbonGroup1});
            this.rbnTabTools.Text = "Tools";
            // 
            // kryptonRibbonGroup1
            // 
            this.kryptonRibbonGroup1.Image = global::DU_Industry_Tool.Properties.Resources.stock_add_bookmark;
            this.kryptonRibbonGroup1.Items.AddRange(new Krypton.Ribbon.KryptonRibbonGroupContainer[] {
            this.kryptonRibbonGroupTriple1});
            this.kryptonRibbonGroup1.MinimumWidth = 100;
            this.kryptonRibbonGroup1.TextLine1 = "Production List";
            // 
            // kryptonRibbonGroupTriple1
            // 
            this.kryptonRibbonGroupTriple1.Items.AddRange(new Krypton.Ribbon.KryptonRibbonGroupItem[] {
            this.RbnBtnProductionList});
            // 
            // RbnBtnProductionList
            // 
            this.RbnBtnProductionList.ImageLarge = global::DU_Industry_Tool.Properties.Resources.stock_add_bookmark;
            this.RbnBtnProductionList.ImageSmall = global::DU_Industry_Tool.Properties.Resources.stock_add_bookmark;
            this.RbnBtnProductionList.TextLine1 = "Production";
            this.RbnBtnProductionList.TextLine2 = "List";
            // 
            // rbnTabThemes
            // 
            this.rbnTabThemes.Groups.AddRange(new Krypton.Ribbon.KryptonRibbonGroup[] {
            this.rbnGrpThemes});
            this.rbnTabThemes.Text = "Themes";
            // 
            // rbnGrpThemes
            // 
            this.rbnGrpThemes.AllowCollapsed = false;
            this.rbnGrpThemes.Image = null;
            this.rbnGrpThemes.Items.AddRange(new Krypton.Ribbon.KryptonRibbonGroupContainer[] {
            this.rbnTrippleThemes});
            this.rbnGrpThemes.TextLine1 = "Themes";
            // 
            // rbnTrippleThemes
            // 
            this.rbnTrippleThemes.Items.AddRange(new Krypton.Ribbon.KryptonRibbonGroupItem[] {
            this.rbnBtnThemeBlue,
            this.rbnBtnThemeSilver,
            this.rbnBtnThemeBlack});
            this.rbnTrippleThemes.MinimumSize = Krypton.Ribbon.GroupItemSize.Large;
            // 
            // rbnBtnThemeBlue
            // 
            this.rbnBtnThemeBlue.ImageLarge = global::DU_Industry_Tool.Properties.Resources.RB2010BlueN;
            this.rbnBtnThemeBlue.ImageSmall = global::DU_Industry_Tool.Properties.Resources.RB2010BlueN;
            this.rbnBtnThemeBlue.TextLine1 = "Office 2010 - Blue";
            this.rbnBtnThemeBlue.Click += new System.EventHandler(this.RbOffice2010Blue_Click);
            // 
            // rbnBtnThemeSilver
            // 
            this.rbnBtnThemeSilver.ImageLarge = global::DU_Industry_Tool.Properties.Resources.Grey_Ball;
            this.rbnBtnThemeSilver.ImageSmall = global::DU_Industry_Tool.Properties.Resources.Grey_Ball;
            this.rbnBtnThemeSilver.KeyTip = "S";
            this.rbnBtnThemeSilver.TextLine1 = "Office 2010 - Silver";
            this.rbnBtnThemeSilver.Click += new System.EventHandler(this.RbOffice2010BSilver_Click);
            // 
            // rbnBtnThemeBlack
            // 
            this.rbnBtnThemeBlack.ImageLarge = global::DU_Industry_Tool.Properties.Resources.Black_Ball;
            this.rbnBtnThemeBlack.ImageSmall = global::DU_Industry_Tool.Properties.Resources.Black_Ball;
            this.rbnBtnThemeBlack.KeyTip = "A";
            this.rbnBtnThemeBlack.TextLine1 = "Office 2010 - Black";
            this.rbnBtnThemeBlack.Click += new System.EventHandler(this.RbOffice2010Black_Click);
            // 
            // kryptonRibbonTab2
            // 
            this.kryptonRibbonTab2.Groups.AddRange(new Krypton.Ribbon.KryptonRibbonGroup[] {
            this.kryptonRibbonGroup5});
            this.kryptonRibbonTab2.KeyTip = "O";
            this.kryptonRibbonTab2.Text = "About";
            // 
            // kryptonRibbonGroup5
            // 
            this.kryptonRibbonGroup5.Image = null;
            this.kryptonRibbonGroup5.Items.AddRange(new Krypton.Ribbon.KryptonRibbonGroupContainer[] {
            this.kryptonRibbonGroupLines1});
            this.kryptonRibbonGroup5.KeyTipDialogLauncher = "OA";
            this.kryptonRibbonGroup5.KeyTipGroup = "A";
            this.kryptonRibbonGroup5.TextLine1 = "About";
            // 
            // kryptonRibbonGroupLines1
            // 
            this.kryptonRibbonGroupLines1.Items.AddRange(new Krypton.Ribbon.KryptonRibbonGroupItem[] {
            this.kryptonRibbonGroupButton19});
            this.kryptonRibbonGroupLines1.MinimumSize = Krypton.Ribbon.GroupItemSize.Large;
            // 
            // kryptonRibbonGroupButton19
            // 
            this.kryptonRibbonGroupButton19.ImageLarge = ((System.Drawing.Image)(resources.GetObject("kryptonRibbonGroupButton19.ImageLarge")));
            this.kryptonRibbonGroupButton19.ImageSmall = ((System.Drawing.Image)(resources.GetObject("kryptonRibbonGroupButton19.ImageSmall")));
            this.kryptonRibbonGroupButton19.TextLine1 = "About";
            this.kryptonRibbonGroupButton19.Click += new System.EventHandler(this.RibbonButtonAboutClick);
            // 
            // kryptonDockableWorkspace
            // 
            this.kryptonDockableWorkspace.ActivePage = this.kryptonPage1;
            this.kryptonDockableWorkspace.AutoHiddenHost = false;
            this.kryptonDockableWorkspace.CompactFlags = Krypton.Workspace.CompactFlags.AtLeastOneVisibleCell;
            this.kryptonDockableWorkspace.ContainerBackStyle = Krypton.Toolkit.PaletteBackStyle.TabDock;
            this.kryptonDockableWorkspace.Dock = System.Windows.Forms.DockStyle.Left;
            this.kryptonDockableWorkspace.Location = new System.Drawing.Point(0, 136);
            this.kryptonDockableWorkspace.MinimumSize = new System.Drawing.Size(300, 400);
            this.kryptonDockableWorkspace.Name = "kryptonDockableWorkspace";
            this.kryptonDockableWorkspace.PaletteMode = Krypton.Toolkit.PaletteMode.Office2010Blue;
            this.kryptonDockableWorkspace.Root.Children.AddRange(new System.ComponentModel.Component[] {
            this.kryptonWorkspaceCell1});
            this.kryptonDockableWorkspace.Root.UniqueName = "D51970B3EA2C496AD51970B3EA2C496A";
            this.kryptonDockableWorkspace.Root.WorkspaceControl = this.kryptonDockableWorkspace;
            this.kryptonDockableWorkspace.SeparatorStyle = Krypton.Toolkit.SeparatorStyle.HighProfile;
            this.kryptonDockableWorkspace.ShowMaximizeButton = false;
            this.kryptonDockableWorkspace.Size = new System.Drawing.Size(450, 864);
            this.kryptonDockableWorkspace.SplitterWidth = 8;
            this.kryptonDockableWorkspace.TabIndex = 0;
            this.kryptonDockableWorkspace.TabStop = true;
            this.kryptonDockableWorkspace.WorkspaceCellAdding += new System.EventHandler<Krypton.Workspace.WorkspaceCellEventArgs>(this.KryptonDockableWorkspace_WorkspaceCellAdding);
            // 
            // kryptonWorkspaceCell1
            // 
            this.kryptonWorkspaceCell1.AllowPageDrag = false;
            this.kryptonWorkspaceCell1.AllowPageReorder = false;
            this.kryptonWorkspaceCell1.AllowTabFocus = false;
            this.kryptonWorkspaceCell1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.kryptonWorkspaceCell1.Button.ButtonDisplayLogic = Krypton.Navigator.ButtonDisplayLogic.None;
            this.kryptonWorkspaceCell1.Button.CloseButtonAction = Krypton.Navigator.CloseButtonAction.None;
            this.kryptonWorkspaceCell1.Button.CloseButtonDisplay = Krypton.Navigator.ButtonDisplay.Hide;
            this.kryptonWorkspaceCell1.Button.ContextButtonAction = Krypton.Navigator.ContextButtonAction.None;
            this.kryptonWorkspaceCell1.Button.ContextButtonDisplay = Krypton.Navigator.ButtonDisplay.Hide;
            this.kryptonWorkspaceCell1.Button.NextButtonAction = Krypton.Navigator.DirectionButtonAction.None;
            this.kryptonWorkspaceCell1.Button.NextButtonDisplay = Krypton.Navigator.ButtonDisplay.Hide;
            this.kryptonWorkspaceCell1.Button.PreviousButtonAction = Krypton.Navigator.DirectionButtonAction.None;
            this.kryptonWorkspaceCell1.Button.PreviousButtonDisplay = Krypton.Navigator.ButtonDisplay.Hide;
            this.kryptonWorkspaceCell1.MaximumSize = new System.Drawing.Size(420, 2000);
            this.kryptonWorkspaceCell1.Name = "kryptonWorkspaceCell1";
            this.kryptonWorkspaceCell1.NavigatorMode = Krypton.Navigator.NavigatorMode.BarCheckButtonGroupOutside;
            this.kryptonWorkspaceCell1.Pages.AddRange(new Krypton.Navigator.KryptonPage[] {
                this.kryptonPage1});
            this.kryptonWorkspaceCell1.SelectedIndex = 0;
            this.kryptonWorkspaceCell1.UniqueName = "B46823ED744B4A87B46823ED744B4A87";
            // 
            // kryptonPage1
            // 
            this.kryptonPage1.AutoHiddenSlideSize = new System.Drawing.Size(200, 200);
            this.kryptonPage1.Controls.Add(this.treeView);
            this.kryptonPage1.Controls.Add(this.searchPanel);
            this.kryptonPage1.Dock = DockStyle.Fill;
            this.kryptonPage1.Flags = 32;
            this.kryptonPage1.LastVisibleSet = true;
            this.kryptonPage1.MinimumSize = new System.Drawing.Size(400, 450);
            this.kryptonPage1.Name = "kryptonPage1";
            this.kryptonPage1.Size = new System.Drawing.Size(450, 831);
            this.kryptonPage1.Text = "Recipes Explorer";
            this.kryptonPage1.TextDescription = "";
            this.kryptonPage1.TextTitle = "";
            this.kryptonPage1.ToolTipTitle = "";
            this.kryptonPage1.UniqueName = "38D886AD20CD402D38D886AD20CD402D";
            // 
            // searchPanel
            // 
            this.searchPanel.AutoSize = true;
            this.searchPanel.Controls.Add(this.SearchBox);
            this.searchPanel.Controls.Add(this.SearchButton);
            this.searchPanel.Controls.Add(this.QuantityBox);
            this.searchPanel.Controls.Add(this.PreviousButton);
            this.searchPanel.Dock = System.Windows.Forms.DockStyle.Top;
            this.searchPanel.Location = new System.Drawing.Point(0, 0);
            this.searchPanel.Name = "searchPanel";
            this.searchPanel.Size = new System.Drawing.Size(418, 39);
            this.searchPanel.TabIndex = 0;
            // 
            // SearchBox
            // 
            this.SearchBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.SearchBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.SearchBox.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.SearchBox.DropDownWidth = 300;
            this.SearchBox.IntegralHeight = false;
            this.SearchBox.Location = new System.Drawing.Point(4, 6);
            this.SearchBox.Margin = new System.Windows.Forms.Padding(4, 4, 0, 0);
            this.SearchBox.MaxDropDownItems = 20;
            this.SearchBox.MaxLength = 40;
            this.SearchBox.Name = "SearchBox";
            this.SearchBox.Size = new System.Drawing.Size(200, 24);
            this.SearchBox.TabIndex = 0;
            // 
            // SearchButton
            // 
            this.SearchButton.AutoSize = true;
            this.SearchButton.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.SearchButton.Location = new System.Drawing.Point(207, 3);
            this.SearchButton.Name = "SearchButton";
            this.SearchButton.Size = new System.Drawing.Size(80, 33);
            this.SearchButton.TabIndex = 1;
            this.SearchButton.Text = "Search";
            this.SearchButton.Click += new System.EventHandler(this.SearchButton_Click);
            // 
            // QuantityBox
            // 
            this.QuantityBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDown;
            this.QuantityBox.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.QuantityBox.Items.AddRange(new object[] {
            "1",
            "2",
            "5",
            "10",
            "20",
            "50",
            "100",
            "200",
            "250",
            "500",
            "1000",
            "5000",
            "10000"});
            this.QuantityBox.Location = new System.Drawing.Point(294, 4);
            this.QuantityBox.Margin = new System.Windows.Forms.Padding(4, 6, 0, 0);
            this.QuantityBox.MaxDropDownItems = 12;
            this.QuantityBox.MaxLength = 6;
            this.QuantityBox.Name = "QuantityBox";
            this.QuantityBox.Size = new System.Drawing.Size(80, 30);
            this.QuantityBox.TabIndex = 2;
            // 
            // PreviousButton
            // 
            this.PreviousButton.AutoSize = true;
            this.PreviousButton.Enabled = false;
            this.PreviousButton.Font = new System.Drawing.Font("Segoe UI", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PreviousButton.Location = new System.Drawing.Point(376, 3);
            this.PreviousButton.Name = "PreviousButton";
            this.PreviousButton.Size = new System.Drawing.Size(40, 33);
            this.PreviousButton.TabIndex = 3;
            this.PreviousButton.Text = "<<";
            this.PreviousButton.Click += new System.EventHandler(this.PreviousButton_Click);
            // 
            // treeView
            // 
            //this.treeView.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            //| System.Windows.Forms.AnchorStyles.Left) 
            //| System.Windows.Forms.AnchorStyles.Right)));
            this.treeView.Dock = DockStyle.Fill;
            this.treeView.Location = new System.Drawing.Point(0, 40);
            this.treeView.MinimumSize = new System.Drawing.Size(400, 400);
            this.treeView.Name = "treeView";
            this.treeView.Size = new System.Drawing.Size(413, 1009);
            this.treeView.TabIndex = 3;
            // 
            // kryptonNavigator1
            // 
            this.kryptonNavigator1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.kryptonNavigator1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.kryptonNavigator1.Bar.BarMultiline = Krypton.Navigator.BarMultiline.Multiline;
            this.kryptonNavigator1.Bar.TabBorderStyle = Krypton.Toolkit.TabBorderStyle.RoundedEqualMedium;
            this.kryptonNavigator1.Location = new System.Drawing.Point(422, 62);
            this.kryptonNavigator1.Name = "kryptonNavigator1";
            this.kryptonNavigator1.Size = new System.Drawing.Size(1422, 930);
            this.kryptonNavigator1.StateCommon.CheckButton.Content.Image.ImageH = Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.kryptonNavigator1.StateCommon.CheckButton.Content.Image.ImageV = Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.kryptonNavigator1.StateCommon.CheckButton.Content.ShortText.TextH = Krypton.Toolkit.PaletteRelativeAlign.Center;
            this.kryptonNavigator1.StateCommon.CheckButton.Content.ShortText.TextV = Krypton.Toolkit.PaletteRelativeAlign.Far;
            this.kryptonNavigator1.TabIndex = 7;
            // 
            // imageListSmall
            // 
            this.imageListSmall.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListSmall.ImageStream")));
            this.imageListSmall.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListSmall.Images.SetKeyName(0, "document_plain.png");
            this.imageListSmall.Images.SetKeyName(1, "preferences.png");
            this.imageListSmall.Images.SetKeyName(2, "information2.png");
            // 
            // kryptonManager
            // 
            this.kryptonManager.GlobalPaletteMode = Krypton.Toolkit.PaletteModeManager.Office2010Blue;
            // 
            // kryptonRibbonGroupGallery1
            // 
            this.kryptonRibbonGroupGallery1.ImageList = null;
            // 
            // MainForm
            // 
            this.AcceptButton = this.SearchButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(120F, 120F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.ClientSize = new System.Drawing.Size(1850, 1000);
            this.Controls.Add(this.kryptonNavigator1);
            this.Controls.Add(this.kryptonDockableWorkspace);
            this.Controls.Add(this.kryptonRibbon);
            this.Font = new System.Drawing.Font("Segoe UI", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(800, 600);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "DU Industry Tool (Mercury)";
            this.Load += new System.EventHandler(this.MainForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonRibbon)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonDockableWorkspace)).EndInit();
            this.kryptonDockableWorkspace.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonPage1)).EndInit();
            this.kryptonPage1.ResumeLayout(false);
            this.kryptonPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.searchPanel)).EndInit();
            this.searchPanel.ResumeLayout(false);
            this.searchPanel.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.SearchBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.kryptonWorkspaceCell1)).EndInit();
            this.kryptonWorkspaceCell1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.kryptonNavigator1)).EndInit();
            this.kryptonNavigator1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ImageList imageListSmall;
        private Krypton.Ribbon.KryptonRibbon kryptonRibbon;
        private Krypton.Ribbon.KryptonRibbonTab kryptonRibbonTab2;
        private Krypton.Ribbon.KryptonRibbonGroup kryptonRibbonGroup5;
        private Krypton.Ribbon.KryptonRibbonGroupLines kryptonRibbonGroupLines1;
        private Krypton.Ribbon.KryptonRibbonGroupButton kryptonRibbonGroupButton19;
        private Krypton.Toolkit.KryptonContextMenuItem ribbonAppButtonExit;
        private Krypton.Toolkit.KryptonContextMenuItem buttonOreValues;
        private Krypton.Toolkit.KryptonContextMenuItem buttonSchematicValues;
        private Krypton.Toolkit.KryptonContextMenuItem buttonSkillLevels;
        private Krypton.Toolkit.KryptonContextMenuItem buttonUpdateMarketValues;
        private Krypton.Toolkit.KryptonContextMenuItem buttonFilterToMarket;
        private Krypton.Toolkit.KryptonContextMenuItem buttonExportToSpreadsheet;
        private Krypton.Toolkit.KryptonContextMenuItem buttonFactoryBreakdownForSelected;
        private Krypton.Toolkit.KryptonManager kryptonManager;
        private Krypton.Navigator.KryptonNavigator kryptonNavigator1;
        private Krypton.Docking.KryptonDockingManager kryptonDockingManager;
        private Krypton.Docking.KryptonDockableWorkspace kryptonDockableWorkspace;
        private Krypton.Workspace.KryptonWorkspaceCell kryptonWorkspaceCell1;
        private Krypton.Navigator.KryptonPage kryptonPage1;
        private System.Windows.Forms.TreeView treeView;
        private Krypton.Toolkit.KryptonPanel searchPanel;
        private Krypton.Toolkit.KryptonComboBox SearchBox;
        private Krypton.Toolkit.KryptonButton SearchButton;
        private Krypton.Toolkit.KryptonButton PreviousButton;
        private Krypton.Toolkit.KryptonComboBox QuantityBox;
        private Krypton.Ribbon.KryptonRibbonTab rbnTabTools;
        private Krypton.Ribbon.KryptonRibbonGroup kryptonRibbonGroup1;
        private Krypton.Ribbon.KryptonRibbonGroupGallery kryptonRibbonGroupGallery1;
        private Krypton.Ribbon.KryptonRibbonGroupTriple kryptonRibbonGroupTriple1;
        private Krypton.Ribbon.KryptonRibbonGroupButton RbnBtnProductionList;
        private Krypton.Ribbon.KryptonRibbonTab rbnTabThemes;
        private Krypton.Ribbon.KryptonRibbonGroup rbnGrpThemes;
        private Krypton.Ribbon.KryptonRibbonGroupTriple rbnTrippleThemes;
        private Krypton.Ribbon.KryptonRibbonGroupButton rbnBtnThemeBlue;
        private Krypton.Ribbon.KryptonRibbonGroupButton rbnBtnThemeSilver;
        private Krypton.Ribbon.KryptonRibbonGroupButton rbnBtnThemeBlack;
        private Krypton.Toolkit.KryptonCommand cmdThemeBlue;
    }
}