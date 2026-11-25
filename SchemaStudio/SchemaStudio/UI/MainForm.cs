using System.Drawing.Drawing2D;
using SchemaStudio.Models;
using SchemaStudio.Services;

namespace SchemaStudio.UI;

public sealed class MainForm : Form
{
    private readonly TreeView tree = new()
    {
        Dock = DockStyle.Fill,
        BorderStyle = BorderStyle.None,
        HideSelection = false,
        FullRowSelect = true,
        ItemHeight = 30,
        ShowLines = false,
        ShowPlusMinus = false,
        ShowRootLines = false,
        DrawMode = TreeViewDrawMode.OwnerDrawAll,
        BackColor = StudioTheme.Background,
        ForeColor = StudioTheme.Text
    };
    private readonly StudioSearchBox search = new("搜索配置表...");
    private readonly StudioTabStrip tabStrip = new() { Dock = DockStyle.Top };
    private readonly Panel tabHost = new() { Dock = DockStyle.Fill, BackColor = StudioTheme.Surface };
    private readonly ColumnInspector inspector = new();
    private readonly Dictionary<string, EditorDocument> opened = new(StringComparer.OrdinalIgnoreCase);
    private readonly ToolTip tips = new() { InitialDelay = 350, ShowAlways = true };
    private readonly Label statusPath = new() { AutoSize = false, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, ForeColor = StudioTheme.Muted, AutoEllipsis = true, Padding = new Padding(4, 0, 8, 0) };
    private readonly Label statusDetail = new() { AutoSize = true, Dock = DockStyle.Right, TextAlign = ContentAlignment.MiddleRight, ForeColor = StudioTheme.Muted, Padding = new Padding(8, 0, 12, 0) };
    private readonly Label sidebarPath = new() { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, ForeColor = StudioTheme.Muted, AutoEllipsis = true, Font = StudioTheme.UiFont(8.25F), Padding = new Padding(8, 0, 8, 0) };
    private readonly DataGridView problems = new();
    private readonly Label problemTitle = new() { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(12, 0, 0, 0), Text = "问题  ·  点击“校验”检查当前编辑内容", ForeColor = StudioTheme.Muted, Font = StudioTheme.UiFont(8.75F) };
    private readonly StudioButton save, undo, redo, moreBtn;
    private readonly SplitContainer workspace, editorSplit, bottomSplit;
    private readonly Panel welcome = new() { Dock = DockStyle.Fill, BackColor = StudioTheme.Surface };
    private PropertyRegistry registry = new();
    private string? currentDir;
    private bool closing;
    private TreeNode? hoverNode;
    private TableEditorControl? Active => tabStrip.SelectedTab?.Tag as TableEditorControl;

    public MainForm(string? initialPath = null)
    {
        Text = "SchemaStudio";
        ClientSize = new Size(1440, 860);
        MinimumSize = new Size(1050, 650);
        Font = StudioTheme.UiFont(9F);
        AutoScaleMode = AutoScaleMode.Dpi;
        BackColor = StudioTheme.Background;
        ForeColor = StudioTheme.Text;
        StartPosition = FormStartPosition.CenterScreen;
        KeyPreview = true;

        // ── Top DBX-style single-row flat header bar ──
        var headerBar = new Panel { Dock = DockStyle.Top, Height = 44, BackColor = StudioTheme.Chrome, Padding = new Padding(10, 6, 10, 6) };
        headerBar.Paint += (_, e) => { using var p = new Pen(StudioTheme.Border); e.Graphics.DrawLine(p, 0, headerBar.Height - 1, headerBar.Width, headerBar.Height - 1); };

        var leftActions = new FlowLayoutPanel { Dock = DockStyle.Left, AutoSize = true, WrapContents = false, FlowDirection = FlowDirection.LeftToRight, BackColor = Color.Transparent, Padding = Padding.Empty, Margin = Padding.Empty };
        var rightActions = new FlowLayoutPanel { Dock = DockStyle.Right, AutoSize = true, WrapContents = false, FlowDirection = FlowDirection.LeftToRight, BackColor = Color.Transparent, Padding = Padding.Empty, Margin = Padding.Empty };

        leftActions.Controls.Add(StudioTheme.Button("打开项目", (_, _) => ChooseDirectory(), icon: StudioIcon.FolderOpen, borderless: true));
        leftActions.Controls.Add(StudioTheme.Button("打开表", (_, _) => ChooseFile(), icon: StudioIcon.FileOpen, borderless: true));
        leftActions.Controls.Add(StudioTheme.Button("新建表", (_, _) => Run(CreateNewTable), icon: StudioIcon.TableAdd, borderless: true));
        leftActions.Controls.Add(MakeDivider());
        save = StudioTheme.Button("保存", (_, _) => SaveActive(), primary: true, icon: StudioIcon.Save);
        tips.SetToolTip(save, "保存数据和字段配置 (Ctrl+S)");
        leftActions.Controls.Add(save);
        leftActions.Controls.Add(StudioTheme.Button("校验", (_, _) => ValidateActive(), icon: StudioIcon.Check, borderless: true));
        leftActions.Controls.Add(StudioTheme.Button("表属性", (_, _) => WithActive(EditTableProperties), icon: StudioIcon.Sliders, borderless: true));
        leftActions.Controls.Add(StudioTheme.Button("属性库", (_, _) => EditRegistry(), icon: StudioIcon.Tag, borderless: true));
        moreBtn = StudioTheme.Button("更多", (_, _) => ShowMoreMenu(), borderless: true);
        moreBtn.ShowDropdown = true;
        leftActions.Controls.Add(moreBtn);

        undo = StudioTheme.IconButton(StudioIcon.Undo, "撤销 (Ctrl+Z)", (_, _) => Undo(), tips);
        redo = StudioTheme.IconButton(StudioIcon.Redo, "重做 (Ctrl+Y)", (_, _) => Redo(), tips);
        rightActions.Controls.Add(undo);
        rightActions.Controls.Add(redo);
        rightActions.Controls.Add(MakeDivider());
        rightActions.Controls.Add(StudioTheme.IconButton(StudioIcon.Refresh, "重新加载当前文件", (_, _) => Run(ReloadActive), tips));
        rightActions.Controls.Add(StudioTheme.IconButton(StudioIcon.External, "用系统默认表格程序打开", (_, _) => WithActive(v => System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(v.Document.Path) { UseShellExecute = true })), tips));

        headerBar.Controls.Add(leftActions);
        headerBar.Controls.Add(rightActions);

        // ── Workspace Split ──
        workspace = new SplitContainer { Dock = DockStyle.Fill, FixedPanel = FixedPanel.Panel1, SplitterWidth = 1, BackColor = StudioTheme.Border };
        BuildSidebar();

        // ── Center + Right Inspector Split ──
        editorSplit = new SplitContainer { Dock = DockStyle.Fill, FixedPanel = FixedPanel.Panel2, SplitterWidth = 1, BackColor = StudioTheme.Border };
        bottomSplit = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, FixedPanel = FixedPanel.Panel2, SplitterWidth = 1, BackColor = StudioTheme.Border };

        var docArea = new Panel { Dock = DockStyle.Fill, BackColor = StudioTheme.Surface };
        tabHost.Controls.Add(welcome);
        docArea.Controls.Add(tabHost);
        docArea.Controls.Add(tabStrip);
        BuildWelcome();

        bottomSplit.Panel1.BackColor = StudioTheme.Surface;
        bottomSplit.Panel1.Controls.Add(docArea);
        editorSplit.Panel1.Controls.Add(bottomSplit);
        editorSplit.Panel2.BackColor = StudioTheme.Surface;
        editorSplit.Panel2.Controls.Add(inspector);
        workspace.Panel2.Controls.Add(editorSplit);
        SetupProblems();

        tabStrip.Selecting += (_, e) => { if (!CommitActive()) e.Cancel = true; };
        tabStrip.SelectedIndexChanged += (_, _) =>
        {
            SyncVisibleTab();
            problems.DataSource = null;
            problemTitle.Text = "问题  ·  点击“校验”检查当前编辑内容";
            RefreshChrome(true);
        };
        tabStrip.CloseRequested += CloseTab;

        // ── Bottom Status Strip (flat clean bar) ──
        var statusBar = new Panel { Dock = DockStyle.Bottom, Height = 26, BackColor = StudioTheme.Background, Padding = new Padding(8, 0, 4, 0) };
        statusBar.Paint += (_, e) => { using var p = new Pen(StudioTheme.Border); e.Graphics.DrawLine(p, 0, 0, statusBar.Width, 0); };
        statusBar.Controls.Add(statusPath);
        statusBar.Controls.Add(statusDetail);

        Controls.Add(workspace);
        Controls.Add(statusBar);
        Controls.Add(headerBar);

        AllowDrop = true;
        DragEnter += (_, e) => { if (e.Data?.GetDataPresent(DataFormats.FileDrop) == true) e.Effect = DragDropEffects.Copy; };
        DragDrop += (_, e) => { if (e.Data?.GetData(DataFormats.FileDrop) is string[] paths) foreach (var path in paths) Run(() => { if (Directory.Exists(path)) OpenDirectory(path); else OpenFile(path); }); };
        FormClosing += OnClosing;
        HandleCreated += (_, _) => StudioTheme.ApplyModernTitleBar(this);
        Shown += (_, _) =>
        {
            StudioTheme.ApplyModernTitleBar(this);
            workspace.Panel1MinSize = 180; workspace.Panel2MinSize = 740; workspace.SplitterDistance = 236;
            editorSplit.Panel1MinSize = 440; editorSplit.Panel2MinSize = 260; editorSplit.SplitterDistance = editorSplit.Width - 292;
            bottomSplit.Panel1MinSize = 240; bottomSplit.Panel2MinSize = 100; bottomSplit.SplitterDistance = bottomSplit.Height - 148;
            if (initialPath != null) Run(() => { if (Directory.Exists(initialPath)) OpenDirectory(initialPath); else { OpenDirectory(Path.GetDirectoryName(Path.GetFullPath(initialPath))!); OpenFile(initialPath); } });
        };
        RefreshChrome(true);
    }

    private static Control MakeDivider() => new Panel { Width = 1, Height = 18, BackColor = StudioTheme.Border, Margin = new Padding(6, 6, 8, 6) };

    private void BuildSidebar()
    {
        workspace.Panel1.BackColor = StudioTheme.Background;

        var sideTop = new Panel { Dock = DockStyle.Top, Height = 38, BackColor = StudioTheme.Background, Padding = new Padding(12, 4, 6, 4) };
        sideTop.Paint += (_, e) => { using var p = new Pen(StudioTheme.Border); e.Graphics.DrawLine(p, 0, sideTop.Height - 1, sideTop.Width, sideTop.Height - 1); };
        var title = new Label { Text = "工作区", Dock = DockStyle.Left, AutoSize = true, TextAlign = ContentAlignment.MiddleLeft, Font = StudioTheme.UiFont(9F, FontStyle.Bold), ForeColor = StudioTheme.Text, Padding = new Padding(0, 7, 0, 0) };
        var sideTools = new FlowLayoutPanel { Dock = DockStyle.Right, AutoSize = true, WrapContents = false, BackColor = Color.Transparent, Padding = Padding.Empty, Margin = Padding.Empty };
        sideTools.Controls.Add(StudioTheme.IconButton(StudioIcon.Refresh, "刷新目录树", (_, _) => Run(BuildTree), tips));
        sideTools.Controls.Add(StudioTheme.IconButton(StudioIcon.FolderOpen, "切换项目目录", (_, _) => ChooseDirectory(), tips));
        sideTop.Controls.Add(title);
        sideTop.Controls.Add(sideTools);

        var searchWrap = new Panel { Dock = DockStyle.Top, Height = 44, BackColor = StudioTheme.Background, Padding = new Padding(10, 6, 10, 6) };
        search.Dock = DockStyle.Fill;
        search.QueryChanged += (_, _) => Run(BuildTree);
        searchWrap.Controls.Add(search);

        var sideFooter = new Panel { Dock = DockStyle.Bottom, Height = 26, BackColor = StudioTheme.Background };
        sideFooter.Paint += (_, e) => { using var p = new Pen(StudioTheme.Border); e.Graphics.DrawLine(p, 0, 0, sideFooter.Width, 0); };
        sideFooter.Controls.Add(sidebarPath);

        StudioTheme.EnableDoubleBuffering(tree);
        tree.DrawNode += DrawTreeNode;
        tree.MouseMove += (_, e) =>
        {
            var n = tree.GetNodeAt(e.Location);
            if (n != hoverNode) { hoverNode = n; tree.Invalidate(); }
        };
        tree.MouseLeave += (_, _) => { if (hoverNode != null) { hoverNode = null; tree.Invalidate(); } };
        tree.NodeMouseClick += (_, e) =>
        {
            if (e.Button != MouseButtons.Left || e.Node == null) return;
            int chevronRight = 10 + e.Node.Level * 16 + 18;
            if (e.X <= chevronRight && e.Node.Nodes.Count > 0)
            {
                if (e.Node.IsExpanded) e.Node.Collapse(); else e.Node.Expand();
            }
        };
        tree.NodeMouseDoubleClick += (_, e) => OpenNode(e.Node);
        tree.KeyDown += (_, e) => { if (e.KeyCode == Keys.Enter) { OpenNode(tree.SelectedNode); e.Handled = true; } };
        tree.BeforeExpand += (_, e) => { if (e.Node?.Tag is NodeTag { IsExcelFile: true } tag) PopulateSheets(e.Node, tag.Path); };

        workspace.Panel1.Controls.Add(tree);
        workspace.Panel1.Controls.Add(searchWrap);
        workspace.Panel1.Controls.Add(sideTop);
        workspace.Panel1.Controls.Add(sideFooter);
    }

    private void DrawTreeNode(object? sender, DrawTreeNodeEventArgs e)
    {
        if (e.Node == null || e.Bounds.Height <= 0) return;
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        var row = new Rectangle(0, e.Bounds.Top, tree.ClientSize.Width, e.Bounds.Height);
        using (var bg = new SolidBrush(StudioTheme.Background)) g.FillRectangle(bg, row);

        bool selected = (e.State & TreeNodeStates.Selected) != 0;
        bool hovered = ReferenceEquals(e.Node, hoverNode);
        if (selected || hovered)
        {
            var pill = new Rectangle(6, row.Top + 2, Math.Max(10, row.Width - 12), row.Height - 4);
            using var path = StudioTheme.RoundedRect(pill, 6);
            using var fill = new SolidBrush(selected ? StudioTheme.ActiveTab : StudioTheme.Hover);
            g.FillPath(fill, path);
        }

        int x = 12 + e.Node.Level * 16;
        int cy = row.Top + row.Height / 2;

        // Crisp chevron for expandable nodes
        if (e.Node.Nodes.Count > 0)
        {
            using var pen = new Pen(StudioTheme.Muted, 1.35f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round };
            if (e.Node.IsExpanded)
                g.DrawLines(pen, new[] { new PointF(x + 2, cy - 2), new PointF(x + 6, cy + 2), new PointF(x + 10, cy - 2) });
            else
                g.DrawLines(pen, new[] { new PointF(x + 4, cy - 4), new PointF(x + 8, cy), new PointF(x + 4, cy + 4) });
        }
        x += 16;

        // Node icon
        var (icon, iconColor) = e.Node.Tag switch
        {
            NodeTag { IsExcelSheet: true } => (StudioIcon.Table, StudioTheme.SheetGreen),
            NodeTag { IsExcelFile: true } => (StudioIcon.Workbook, StudioTheme.SheetGreen),
            NodeTag { IsCsv: true } => (StudioIcon.Csv, StudioTheme.CsvBlue),
            _ => (StudioIcon.Folder, StudioTheme.FolderTint)
        };
        StudioTheme.DrawIcon(g, icon, new Rectangle(x, cy - 7, 14, 14), iconColor);
        x += 20;

        var textRect = new Rectangle(x, row.Top, Math.Max(10, row.Right - x - 10), row.Height);
        var font = e.Node.Level == 0 ? StudioTheme.UiFont(9F, FontStyle.Bold) : Font;
        TextRenderer.DrawText(g, e.Node.Text, font, textRect, selected ? StudioTheme.Text : StudioTheme.Text,
            TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
    }

    private void BuildWelcome()
    {
        var center = new TableLayoutPanel
        {
            AutoSize = true,
            ColumnCount = 1,
            BackColor = StudioTheme.Surface,
            Anchor = AnchorStyles.None
        };
        var iconHost = new Panel { Size = new Size(48, 48), Margin = new Padding(0, 0, 0, 8), Anchor = AnchorStyles.None };
        iconHost.Paint += (_, e) => StudioTheme.DrawIcon(e.Graphics, StudioIcon.Table, new Rectangle(6, 6, 36, 36), StudioTheme.SubtleText);
        var title = new Label { Text = "SchemaStudio 配置工作台", AutoSize = true, Font = StudioTheme.UiFont(14F, FontStyle.Bold), ForeColor = StudioTheme.Text, Anchor = AnchorStyles.None, Margin = new Padding(0, 0, 0, 6) };
        var sub = new Label { Text = "打开 CSV 或 Excel 配置表，同步查看数据与管理 .schema.json 字段快照。", AutoSize = true, ForeColor = StudioTheme.Muted, Anchor = AnchorStyles.None, Margin = new Padding(0, 0, 0, 18) };
        var actions = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Anchor = AnchorStyles.None };
        actions.Controls.Add(StudioTheme.Button("打开项目目录", (_, _) => ChooseDirectory(), primary: true, icon: StudioIcon.FolderOpen));
        actions.Controls.Add(StudioTheme.Button("打开单个表", (_, _) => ChooseFile(), icon: StudioIcon.FileOpen));
        actions.Controls.Add(StudioTheme.Button("新建表", (_, _) => Run(CreateNewTable), icon: StudioIcon.TableAdd));
        var hint = new Label { Text = "快捷键：Ctrl+O 打开  ·  Ctrl+S 保存  ·  F5 校验  ·  支持拖拽文件到窗口", AutoSize = true, Font = StudioTheme.UiFont(8.5F), ForeColor = StudioTheme.SubtleText, Anchor = AnchorStyles.None, Margin = new Padding(0, 16, 0, 0) };

        center.Controls.Add(iconHost);
        center.Controls.Add(title);
        center.Controls.Add(sub);
        center.Controls.Add(actions);
        center.Controls.Add(hint);

        var host = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 1, BackColor = StudioTheme.Surface };
        host.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        host.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        host.Controls.Add(center, 0, 0);
        welcome.Controls.Add(host);
    }

    private void SetupProblems()
    {
        var header = new Panel { Dock = DockStyle.Top, Height = 32, BackColor = StudioTheme.Background };
        header.Paint += (_, e) => { using var p = new Pen(StudioTheme.Border); e.Graphics.DrawLine(p, 0, header.Height - 1, header.Width, header.Height - 1); };
        header.Controls.Add(problemTitle);

        problems.Dock = DockStyle.Fill;
        StudioTheme.Grid(problems);
        problems.ReadOnly = true;
        problems.RowHeadersVisible = false;
        problems.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        problems.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "行", DataPropertyName = nameof(ValidationError.Row), Width = 65 });
        problems.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "字段", DataPropertyName = nameof(ValidationError.Column), Width = 150 });
        problems.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "问题说明", DataPropertyName = nameof(ValidationError.Message), AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
        problems.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "实际值", DataPropertyName = nameof(ValidationError.Value), Width = 160 });
        problems.CellDoubleClick += (_, e) =>
        {
            if (e.RowIndex < 0 || Active == null || problems.Rows[e.RowIndex].DataBoundItem is not ValidationError error) return;
            var grid = Active.DataGrid;
            int col = Active.Sheet.Columns.FindIndex(c => c.ColName == error.Column);
            if (error.Row >= 2 && error.Row - 2 < grid.RowCount) { grid.CurrentCell = grid[Math.Max(0, col), error.Row - 2]; grid.Focus(); }
        };
        bottomSplit.Panel2.BackColor = StudioTheme.Surface;
        bottomSplit.Panel2.Controls.Add(problems);
        bottomSplit.Panel2.Controls.Add(header);
    }

    private void SyncVisibleTab()
    {
        var active = Active;
        foreach (Control c in tabHost.Controls)
        {
            if (c is TableEditorControl) c.Visible = ReferenceEquals(c, active);
        }
        welcome.Visible = active == null;
        tabStrip.Visible = tabStrip.Count > 0;
        active?.BringToFront();
    }

    public void OpenDirectory(string path)
    {
        if (!CloseAll()) return;
        var loadedRegistry = DataIO.LoadRegistry(Path.Combine(path, "schema.registry.json"));
        currentDir = Path.GetFullPath(path);
        registry = loadedRegistry;
        if (!registry.Properties.Any(p => p.Key.Equals("cs_type", StringComparison.OrdinalIgnoreCase)))
            registry.Properties.Add(new PropertyDef { Key = "cs_type", Title = "终端类型", Scope = "both", Kind = "enum", EnumValues = new() { "ALL", "C", "S" }, Default = "ALL" });
        sidebarPath.Text = currentDir;
        tips.SetToolTip(sidebarPath, currentDir);
        BuildTree();
        RefreshChrome(true);
    }

    public void OpenFile(string path, string? sheetName = null)
    {
        if (!CommitActive()) return;
        path = Path.GetFullPath(path);
        if (!opened.TryGetValue(path, out var document))
        {
            document = DocumentStore.Open(path);
            opened.Add(path, document);
            document.Changed += DocumentChanged;
        }
        var sheet = sheetName == null ? document.Sheets[0] : document.Sheets.First(s => s.Name == sheetName);
        foreach (var existing in tabStrip.Tabs)
        {
            if (existing.Tag is TableEditorControl editor && editor.Document == document && editor.Sheet == sheet)
            {
                tabStrip.SelectTab(existing);
                return;
            }
        }
        var view = new TableEditorControl(document, sheet, registry) { Dock = DockStyle.Fill };
        view.ColumnSelected += (_, _) => { if (Active == view) inspector.ShowColumn(view, registry); };
        view.AddRowRequested += (_, _) => WithActive(v => v.Document.AddRow(v.Sheet));
        view.AddColumnRequested += (_, _) => WithActive(AddColumn);
        view.DeleteMenuRequested += (_, _) => ShowDeleteMenu();

        tabHost.Controls.Add(view);
        var tab = new StudioTabItem(view, FormatTabTitle(view))
        {
            IsDirty = document.IsDirty,
            Icon = document.IsCsv ? StudioIcon.Csv : StudioIcon.Table,
            IconColor = document.IsCsv ? StudioTheme.CsvBlue : StudioTheme.SheetGreen
        };
        tabStrip.AddTab(tab);
        SyncVisibleTab();
        RefreshChrome(true);
    }

    private static string FormatTabTitle(TableEditorControl view)
        => view.Document.IsCsv ? Path.GetFileName(view.Document.Path) : $"{Path.GetFileName(view.Document.Path)} / {view.Sheet.Name}";

    private void ChooseDirectory()
    {
        using var dialog = new FolderBrowserDialog { Description = "选择配置表根目录", SelectedPath = currentDir ?? "" };
        if (dialog.ShowDialog(this) == DialogResult.OK) Run(() => OpenDirectory(dialog.SelectedPath));
    }

    private void ChooseFile()
    {
        using var dialog = new OpenFileDialog { Filter = "配置表|*.csv;*.xlsx;*.xls", Multiselect = true, InitialDirectory = currentDir ?? "" };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        if (currentDir == null) Run(() => OpenDirectory(Path.GetDirectoryName(dialog.FileNames[0])!));
        foreach (var path in dialog.FileNames) Run(() => OpenFile(path));
    }

    private void BuildTree()
    {
        tree.BeginUpdate();
        try
        {
            tree.Nodes.Clear();
            if (currentDir == null) return;
            var root = new TreeNode(Path.GetFileName(currentDir)) { Tag = new NodeTag(currentDir, null) };
            tree.Nodes.Add(root);
            AddDirectory(root, currentDir);
            root.Expand();
        }
        finally { tree.EndUpdate(); }
    }

    private void AddDirectory(TreeNode parent, string directory)
    {
        foreach (string sub in Directory.GetDirectories(directory).Where(d => !Path.GetFileName(d).StartsWith('.') && Path.GetFileName(d) is not "bin" and not "obj"))
        {
            if ((File.GetAttributes(sub) & FileAttributes.ReparsePoint) != 0) continue;
            var child = new TreeNode(Path.GetFileName(sub)) { Tag = new NodeTag(sub, null) };
            AddDirectory(child, sub);
            if (child.Nodes.Count > 0) parent.Nodes.Add(child);
        }
        foreach (string file in Directory.GetFiles(directory).Where(f => UiDefaults.FileExts.Contains(Path.GetExtension(f).ToLowerInvariant())).OrderBy(f => f, StringComparer.OrdinalIgnoreCase))
        {
            string name = Path.GetFileName(file);
            if (name.StartsWith("~$") || !name.Contains(search.Query, StringComparison.OrdinalIgnoreCase)) continue;
            var node = new TreeNode(name) { Tag = new NodeTag(file, null) };
            if (!Path.GetExtension(file).Equals(".csv", StringComparison.OrdinalIgnoreCase)) node.Nodes.Add(new TreeNode("展开工作表…"));
            parent.Nodes.Add(node);
        }
    }

    private void PopulateSheets(TreeNode node, string path)
    {
        if (node.Nodes.Count != 1 || node.Nodes[0].Tag != null) return;
        Run(() =>
        {
            var document = opened.GetValueOrDefault(path) ?? DocumentStore.Open(path);
            node.Nodes.Clear();
            foreach (var sheet in document.Sheets) node.Nodes.Add(new TreeNode(sheet.Name) { Tag = new NodeTag(path, sheet.Name) });
        });
    }

    private void OpenNode(TreeNode? node) { if (node?.Tag is NodeTag tag && !Directory.Exists(tag.Path)) Run(() => OpenFile(tag.Path, tag.Sheet)); }
    private void DocumentChanged(object? sender, EventArgs e) { if (IsHandleCreated && !IsDisposed) BeginInvoke((Action)(() => RefreshChrome(false))); }

    private void RefreshChrome(bool forceInspector)
    {
        foreach (var tab in tabStrip.Tabs)
        {
            if (tab.Tag is TableEditorControl view)
            {
                tab.Title = FormatTabTitle(view);
                tab.IsDirty = view.Document.IsDirty;
            }
        }
        tabStrip.RefreshTabs();
        SyncVisibleTab();

        save.Enabled = Active != null;
        undo.Enabled = Active?.Document.CanUndo == true;
        redo.Enabled = Active?.Document.CanRedo == true;
        inspector.ShowColumn(Active, registry, forceInspector);

        statusPath.Text = Active == null ? (currentDir ?? "就绪 · 打开配置目录开始编辑") : Active.Document.Path + (Active.Document.IsDirty ? "  ·  未保存" : "  ·  已保存");
        statusDetail.Text = Active == null ? "CSV / XLSX / XLS" : $"{Active.Sheet.RowCount:N0} 行  ·  {Active.Sheet.Columns.Count} 列  ·  {Active.Sheet.Columns.Count(c => ColumnConfiguration.Missing(c, registry) != null)} 列待配置";
        Text = "SchemaStudio" + (Active == null ? "" : " - " + Path.GetFileName(Active.Document.Path) + (Active.Document.IsDirty ? " *" : ""));
    }

    private bool CommitActive() => Active == null || Active.CommitEdit() && ValidateChildren();

    private bool SaveDocument(EditorDocument document)
    {
        if (!CommitActive()) return false;
        try
        {
            string backup = DocumentStore.Save(document);
            RefreshChrome(false);
            statusPath.Text = "已保存数据与字段配置 · 备份：" + backup;
            return true;
        }
        catch (Exception error)
        {
            MessageBox.Show(this, error.Message, "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
    }

    private void SaveActive() { if (Active != null) SaveDocument(Active.Document); }
    private void Undo() { if (Active != null && CommitActive()) { Active.Document.Undo(); Active.RefreshModel(); inspector.ShowColumn(Active, registry, true); } }
    private void Redo() { if (Active != null && CommitActive()) { Active.Document.Redo(); Active.RefreshModel(); inspector.ShowColumn(Active, registry, true); } }
    private void WithActive(Action<TableEditorControl> action) { if (Active != null && CommitActive()) Run(() => action(Active)); }

    private void AddColumn(TableEditorControl view)
    {
        using var dialog = new ColumnDefinitionForm(registry, view.Sheet.Columns.Select(c => c.ColName ?? ""));
        if (dialog.ShowDialog(this) == DialogResult.OK && dialog.Result != null) view.Document.AddColumn(view.Sheet, dialog.Result);
    }

    private void ShowMoreMenu()
    {
        var menu = new ContextMenuStrip { Font = StudioTheme.UiFont(9F) };
        menu.Items.Add("追加行", null, (_, _) => WithActive(v => v.Document.AddRow(v.Sheet))).Enabled = Active != null;
        menu.Items.Add("追加列", null, (_, _) => WithActive(AddColumn)).Enabled = Active != null;
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("删除选中行（CSV）", null, (_, _) => WithActive(v => v.Document.DeleteRow(v.Sheet, v.SelectedDataRow))).Enabled = Active?.Document.IsCsv == true;
        menu.Items.Add("删除选中列（CSV）", null, (_, _) => WithActive(v => v.Document.DeleteColumn(v.Sheet, v.SelectedColumn))).Enabled = Active?.Document.IsCsv == true;
        menu.Items.Add(new ToolStripSeparator());
        menu.Items.Add("全部保存 (Ctrl+Shift+S)", null, (_, _) => { if (CommitActive()) foreach (var doc in opened.Values.Where(d => d.IsDirty)) if (!SaveDocument(doc)) break; }).Enabled = opened.Values.Any(d => d.IsDirty);
        menu.Items.Add("外部打开源文件", null, (_, _) => WithActive(v => System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(v.Document.Path) { UseShellExecute = true }))).Enabled = Active != null;
        menu.Items.Add("重新加载当前文件", null, (_, _) => Run(ReloadActive)).Enabled = Active != null;
        menu.Closed += (_, _) => menu.Dispose();
        menu.Show(moreBtn, new Point(0, moreBtn.Height + 2));
    }

    private void ShowDeleteMenu()
    {
        if (Active == null) return;
        var menu = new ContextMenuStrip { Font = StudioTheme.UiFont(9F) };
        menu.Items.Add("删除选中行（CSV）", null, (_, _) => WithActive(v => v.Document.DeleteRow(v.Sheet, v.SelectedDataRow))).Enabled = Active.Document.IsCsv;
        menu.Items.Add("删除选中列（CSV）", null, (_, _) => WithActive(v => v.Document.DeleteColumn(v.Sheet, v.SelectedColumn))).Enabled = Active.Document.IsCsv;
        menu.Closed += (_, _) => menu.Dispose();
        menu.Show(Cursor.Position);
    }

    private void ValidateActive()
    {
        if (Active == null || !CommitActive()) return;
        var errors = DocumentValidation.Validate(Active.Sheet, registry);
        problems.DataSource = errors;
        int formulas = Active.Sheet.Cells.Values.Count(c => c.IsFormula);
        problemTitle.Text = $"问题  ·  {errors.Count} 处  ·  {Active.Sheet.Name}" + (formulas > 0 ? $"  ·  {formulas} 个公式需由 Excel 重算和校验" : "  ·  已检查当前未保存内容");
    }

    private void EditTableProperties(TableEditorControl view)
    {
        using var dialog = new AssignPropertyForm(registry.Properties.Where(d => d.Scope is "table" or "both").ToList(), view.Sheet.Properties) { Text = "表属性" };
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        var before = view.Sheet.Properties; var after = dialog.Result;
        view.Document.Execute("修改表属性", () => view.Sheet.Properties = after, () => view.Sheet.Properties = before);
    }

    private void EditRegistry()
    {
        if (currentDir == null) { ChooseDirectory(); if (currentDir == null) return; }
        using var dialog = new PropertyLibraryForm(registry);
        if (dialog.ShowDialog(this) == DialogResult.OK) Run(() =>
        {
            DataIO.SaveRegistry(Path.Combine(currentDir, "schema.registry.json"), dialog.Result);
            registry = dialog.Result;
            foreach (var tab in tabStrip.Tabs)
                if (tab.Tag is TableEditorControl view) { view.Registry = registry; view.RefreshModel(); }
            RefreshChrome(true);
        });
    }

    private bool ConfirmClose(EditorDocument document)
    {
        if (!document.IsDirty) return true;
        var answer = MessageBox.Show(this, $"{Path.GetFileName(document.Path)} 有未保存修改，是否保存？", "关闭配置表", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        return answer == DialogResult.No || answer == DialogResult.Yes && SaveDocument(document);
    }

    private void CloseTab(StudioTabItem tab)
    {
        if (!CommitActive() || tab.Tag is not TableEditorControl view) return;
        bool last = tabStrip.Tabs.Count(p => (p.Tag as TableEditorControl)?.Document == view.Document) == 1;
        if (last && !ConfirmClose(view.Document)) return;
        tabStrip.RemoveTab(tab);
        tabHost.Controls.Remove(view);
        view.Dispose();
        if (last) { opened.Remove(view.Document.Path); view.Document.Changed -= DocumentChanged; }
        RefreshChrome(true);
    }

    private bool CloseAll()
    {
        if (!CommitActive()) return false;
        foreach (var document in opened.Values)
        {
            var tab = tabStrip.Tabs.FirstOrDefault(p => (p.Tag as TableEditorControl)?.Document == document);
            if (tab != null) tabStrip.SelectTab(tab);
            if (!ConfirmClose(document)) return false;
        }
        foreach (var tab in tabStrip.Tabs.ToArray())
        {
            tabStrip.RemoveTab(tab);
            if (tab.Content is Control c) { tabHost.Controls.Remove(c); c.Dispose(); }
        }
        foreach (var doc in opened.Values) doc.Changed -= DocumentChanged;
        opened.Clear();
        return true;
    }

    private void OnClosing(object? sender, FormClosingEventArgs e) { if (closing) return; if (!CloseAll()) { e.Cancel = true; return; } closing = true; }
    private void Run(Action action) { try { action(); } catch (Exception error) { MessageBox.Show(this, error.Message, "操作未完成", MessageBoxButtons.OK, MessageBoxIcon.Information); } }

    protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
    {
        switch (keyData)
        {
            case Keys.Control | Keys.S: SaveActive(); return true;
            case Keys.Control | Keys.Shift | Keys.S: if (CommitActive()) foreach (var doc in opened.Values.Where(d => d.IsDirty)) if (!SaveDocument(doc)) break; return true;
            case Keys.Control | Keys.O: ChooseFile(); return true;
            case Keys.Control | Keys.Z: Undo(); return true;
            case Keys.Control | Keys.Y: Redo(); return true;
            case Keys.Control | Keys.W: if (tabStrip.SelectedTab is { } tab) CloseTab(tab); return true;
            case Keys.F5: ValidateActive(); return true;
        }
        return base.ProcessCmdKey(ref msg, keyData);
    }

    private void CreateNewTable()
    {
        if (currentDir == null) { ChooseDirectory(); if (currentDir == null) return; }
        using var dialog = new NewTableForm(currentDir, DataIO.GetAllDirs(currentDir).Where(d => !d.Contains(".schemastudio-backups")).ToList(), registry, UiDefaults.AllowedTypes);
        if (dialog.ShowDialog(this) != DialogResult.OK) return;
        string path = Path.Combine(dialog.ResultDir!, dialog.ResultFileName! + "." + dialog.ResultFormat);
        if (File.Exists(path) || File.Exists(DataIO.GetSchemaJsonPath(path))) throw new IOException("同名数据文件或配置文件已经存在。");
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        var sheets = dialog.ResultSheetColumns;
        if (dialog.ResultFormat == "csv") DataIO.WriteCsvHeader(path, sheets.First().Value.Select(c => c.ColName));
        else if (dialog.ResultFormat == "xlsx") DataIO.TryWriteXlsxMulti(path, sheets.Select(s => (s.Key, s.Value.Select(c => c.ColName).ToArray())), out _);
        else DataIO.TryWriteXlsMulti(path, sheets.Select(s => (s.Key, s.Value.Select(c => c.ColName).ToArray())), out _);
        var document = DocumentStore.Open(path);
        foreach (var sheet in document.Sheets)
        {
            var source = document.IsCsv ? sheets.First().Value : sheets[sheet.Name];
            sheet.Columns = source.Select((c, index) => new ColumnSchema { ColName = c.ColName, ValueType = c.ValueType, PkOrder = c.PkOrder, PrimaryKey = c.PkOrder >= 0, ColSort = index, ColProperties = c.ColProperties, IsConfigured = true }).ToList();
            sheet.Properties = dialog.ResultSheetTableProps.GetValueOrDefault(document.IsCsv ? sheets.First().Key : sheet.Name) ?? new();
        }
        DocumentStore.Save(document);
        BuildTree();
        OpenFile(path);
    }

    private void ReloadActive()
    {
        if (Active == null || !CommitActive()) return;
        var previous = Active.Document;
        string selected = Active.Sheet.Name;
        if (!ConfirmClose(previous)) return;
        var replacement = DocumentStore.Open(previous.Path);
        var tabs = tabStrip.Tabs.Where(p => (p.Tag as TableEditorControl)?.Document == previous).ToList();
        var sheets = tabs.Select(p => ((TableEditorControl)p.Tag!).Sheet.Name).ToList();
        foreach (var tab in tabs)
        {
            tabStrip.RemoveTab(tab);
            if (tab.Content is Control c) { tabHost.Controls.Remove(c); c.Dispose(); }
        }
        previous.Changed -= DocumentChanged;
        replacement.Changed += DocumentChanged;
        opened[previous.Path] = replacement;
        foreach (string name in sheets.Where(n => replacement.Sheets.Any(s => s.Name == n))) OpenFile(replacement.Path, name);
        OpenFile(replacement.Path, replacement.Sheets.Any(s => s.Name == selected) ? selected : replacement.Sheets[0].Name);
        RefreshChrome(true);
    }
}
