using System.Drawing.Drawing2D;
using System.Globalization;
using CsvHelper;
using CsvHelper.Configuration;
using SchemaStudio.Models;
using SchemaStudio.Services;

namespace SchemaStudio.UI;

public sealed class TableEditorControl : UserControl
{
    public EditorDocument Document { get; }
    public EditorSheet Sheet { get; }
    public PropertyRegistry Registry { get; set; }
    public DataGridView DataGrid { get; } = new();
    private readonly DataGridView structure = new();
    private readonly StudioSegmentedControl viewSwitch = new("数据", "结构", "JSON 预览");
    private readonly Panel contentHost = new() { Dock = DockStyle.Fill, BackColor = StudioTheme.Surface };
    private readonly Panel jsonHost = new() { Dock = DockStyle.Fill, BackColor = StudioTheme.Surface, Padding = new Padding(14), Visible = false };
    private readonly TextBox formula = new() { BorderStyle = BorderStyle.None, Dock = DockStyle.Fill, BackColor = StudioTheme.Background, ForeColor = StudioTheme.Text, Font = StudioTheme.UiFont(9F) };
    private readonly Label address = new() { Text = "A2", TextAlign = ContentAlignment.MiddleCenter, Dock = DockStyle.Fill, ForeColor = StudioTheme.Muted, Font = StudioTheme.UiFont(8.5F, FontStyle.Bold) };
    private readonly Label lockState = new() { AutoSize = true, ForeColor = StudioTheme.Muted, Margin = new Padding(4, 5, 4, 0), Font = StudioTheme.UiFont(8.5F) };
    private readonly QuestionButton protectionHelp = new() { Name = "protected_help", Margin = new Padding(0, 1, 2, 0) };
    private readonly Label breadcrumbFile = new() { AutoSize = true, ForeColor = StudioTheme.Muted, Margin = new Padding(0, 6, 4, 0) };
    private readonly Label breadcrumbSep = new() { AutoSize = true, Text = "/", ForeColor = StudioTheme.SubtleText, Margin = new Padding(0, 6, 4, 0) };
    private readonly Label breadcrumbSheet = new() { AutoSize = true, ForeColor = StudioTheme.Text, Font = StudioTheme.UiFont(9F, FontStyle.Bold), Margin = new Padding(0, 6, 8, 0) };
    private readonly Label breadcrumbMeta = new() { AutoSize = true, ForeColor = StudioTheme.Muted, Font = StudioTheme.UiFont(8.5F), Margin = new Padding(0, 7, 6, 0) };
    private readonly Panel formulaBar;
    public string? ProtectionMessage { get; private set; }
    private readonly TextBox preview = new() { Dock = DockStyle.Fill, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Both, WordWrap = false, BorderStyle = BorderStyle.None, BackColor = StudioTheme.Surface, ForeColor = StudioTheme.Text, Font = new Font("Consolas", 10) };
    private readonly Label notice = new() { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, ForeColor = StudioTheme.Muted, Padding = new Padding(12, 0, 8, 0), AutoEllipsis = true, Font = StudioTheme.UiFont(8.5F) };
    private readonly StudioButton addRowBtn, addColBtn, delBtn;
    private bool refreshing;
    private int headerSelection;

    public int SelectedColumn => viewSwitch.SelectedIndex == 1 ? structure.CurrentCell?.RowIndex ?? 0 : DataGrid.CurrentCell?.ColumnIndex ?? headerSelection;
    public int SelectedDataRow => (DataGrid.CurrentCell?.RowIndex ?? 0) + 1;
    public event EventHandler? ColumnSelected;
    public event EventHandler? AddRowRequested;
    public event EventHandler? AddColumnRequested;
    public event EventHandler? DeleteMenuRequested;

    public TableEditorControl(EditorDocument document, EditorSheet sheet, PropertyRegistry? registry = null)
    {
        Document = document; Sheet = sheet; Registry = registry ?? new();
        Dock = DockStyle.Fill; BackColor = StudioTheme.Surface;

        // ── DBX-style Context / Breadcrumb Bar ──
        var contextBar = new Panel { Dock = DockStyle.Top, Height = 40, BackColor = StudioTheme.Surface, Padding = new Padding(12, 4, 10, 4) };
        contextBar.Paint += (_, e) => { using var p = new Pen(StudioTheme.Border); e.Graphics.DrawLine(p, 0, contextBar.Height - 1, contextBar.Width, contextBar.Height - 1); };

        var contextLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, BackColor = StudioTheme.Surface, Margin = Padding.Empty, Padding = Padding.Empty };
        contextLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        contextLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        contextLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

        var crumb = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, BackColor = StudioTheme.Surface, Margin = Padding.Empty, Padding = Padding.Empty };
        var iconBox = new Panel { Size = new Size(16, 30), Margin = new Padding(0, 0, 6, 0), BackColor = StudioTheme.Surface };
        iconBox.Paint += (_, e) => StudioTheme.DrawIcon(e.Graphics, Document.IsCsv ? StudioIcon.Csv : StudioIcon.Table, new Rectangle(0, 7, 15, 15), Document.IsCsv ? StudioTheme.CsvBlue : StudioTheme.SheetGreen);
        breadcrumbFile.Text = Path.GetFileName(Document.Path);
        breadcrumbFile.Margin = new Padding(0, 7, 4, 0);
        breadcrumbSep.Margin = new Padding(0, 7, 4, 0);
        breadcrumbSheet.Text = Document.IsCsv ? "数据表" : Sheet.Name;
        breadcrumbSheet.Margin = new Padding(0, 7, 8, 0);
        breadcrumbMeta.Margin = new Padding(0, 8, 6, 0);
        crumb.Controls.Add(iconBox);
        crumb.Controls.Add(breadcrumbFile);
        crumb.Controls.Add(breadcrumbSep);
        crumb.Controls.Add(breadcrumbSheet);
        crumb.Controls.Add(breadcrumbMeta);

        var rightTools = new FlowLayoutPanel { AutoSize = true, WrapContents = false, BackColor = StudioTheme.Surface, Margin = Padding.Empty, Padding = Padding.Empty, Anchor = AnchorStyles.Right };
        addRowBtn = StudioTheme.Button("追加行", (_, _) => AddRowRequested?.Invoke(this, EventArgs.Empty), icon: StudioIcon.RowAdd, borderless: true);
        addColBtn = StudioTheme.Button("追加列", (_, _) => AddColumnRequested?.Invoke(this, EventArgs.Empty), icon: StudioIcon.ColAdd, borderless: true);
        delBtn = StudioTheme.Button("删除", (_, _) => DeleteMenuRequested?.Invoke(this, EventArgs.Empty), icon: StudioIcon.Trash, borderless: true);
        delBtn.ShowDropdown = true;
        viewSwitch.Margin = new Padding(6, 1, 0, 0);
        viewSwitch.SelectedIndexChanged += (_, _) => SwitchMode();

        rightTools.Controls.Add(addRowBtn);
        rightTools.Controls.Add(addColBtn);
        if (Document.IsCsv) rightTools.Controls.Add(delBtn);
        rightTools.Controls.Add(viewSwitch);

        contextLayout.Controls.Add(crumb, 0, 0);
        contextLayout.Controls.Add(rightTools, 1, 0);
        contextBar.Controls.Add(contextLayout);

        // ── Sleek Formula / Cell Inspector Bar ──
        formulaBar = new Panel { Dock = DockStyle.Top, Height = 34, BackColor = StudioTheme.Background, Padding = new Padding(8, 2, 8, 2) };
        formulaBar.Paint += (_, e) => { using var p = new Pen(StudioTheme.Border); e.Graphics.DrawLine(p, 0, formulaBar.Height - 1, formulaBar.Width, formulaBar.Height - 1); };
        var bar = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 4, RowCount = 1, BackColor = StudioTheme.Background, Margin = Padding.Empty, Padding = Padding.Empty };
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 52));
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 28));
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        bar.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        address.BackColor = StudioTheme.Background;
        address.Margin = Padding.Empty;
        bar.Controls.Add(address, 0, 0);
        bar.Controls.Add(new Label { Text = "fx", ForeColor = StudioTheme.Accent, BackColor = StudioTheme.Background, Dock = DockStyle.Fill, Margin = Padding.Empty, TextAlign = ContentAlignment.MiddleCenter, Font = StudioTheme.UiFont(9F, FontStyle.Italic | FontStyle.Bold) }, 1, 0);
        var formulaWrap = new Panel { Dock = DockStyle.Fill, Margin = Padding.Empty, Padding = new Padding(4, 6, 4, 0), BackColor = StudioTheme.Background };
        formulaWrap.Controls.Add(formula);
        bar.Controls.Add(formulaWrap, 2, 0);
        var state = new FlowLayoutPanel { AutoSize = true, WrapContents = false, FlowDirection = FlowDirection.LeftToRight, Padding = Padding.Empty, Margin = new Padding(0, 2, 0, 0), Dock = DockStyle.Right, BackColor = StudioTheme.Background };
        state.Controls.Add(lockState);
        state.Controls.Add(protectionHelp);
        bar.Controls.Add(state, 3, 0);
        protectionHelp.Click += (_, _) => ShowProtection();
        formulaBar.Controls.Add(bar);

        // ── Content Views ──
        ConfigureDataGrid();
        ConfigureStructureGrid();
        structure.Visible = false;
        jsonHost.Controls.Add(preview);
        contentHost.Controls.Add(DataGrid);
        contentHost.Controls.Add(structure);
        contentHost.Controls.Add(jsonHost);

        // ── Bottom Notice Bar ──
        var noticeBar = new Panel { Dock = DockStyle.Bottom, Height = 28, BackColor = StudioTheme.Background };
        noticeBar.Paint += (_, e) => { using var p = new Pen(StudioTheme.Border); e.Graphics.DrawLine(p, 0, 0, noticeBar.Width, 0); };
        noticeBar.Controls.Add(notice);

        Controls.Add(contentHost);
        Controls.Add(formulaBar);
        Controls.Add(contextBar);
        Controls.Add(noticeBar);

        Document.Changed += DocumentChanged;
        RefreshModel();
    }

    private void SwitchMode()
    {
        int idx = viewSwitch.SelectedIndex;
        DataGrid.Visible = idx == 0;
        formulaBar.Visible = idx == 0;
        structure.Visible = idx == 1;
        jsonHost.Visible = idx == 2;
        if (idx == 2) preview.Text = DocumentStore.SchemaJson(Document);
        ColumnSelected?.Invoke(this, EventArgs.Empty);
    }

    private void ConfigureDataGrid()
    {
        DataGrid.Dock = DockStyle.Fill; StudioTheme.Grid(DataGrid); DataGrid.VirtualMode = true;
        DataGrid.CellValueNeeded += (_, e) => e.Value = Sheet.Cell(e.RowIndex + 1, e.ColumnIndex).Display;
        DataGrid.CellBeginEdit += (_, e) =>
        {
            if (Document.ReadOnlyReason != null || Sheet.ProtectionAt(e.RowIndex + 1, e.ColumnIndex) != null)
            { e.Cancel = true; UpdateFormulaBar(); return; }
        };
        DataGrid.EditingControlShowing += (_, e) =>
        {
            if (e.Control is TextBox input && DataGrid.CurrentCell is { } cell)
                input.Text = Sheet.Cell(cell.RowIndex + 1, cell.ColumnIndex).Text;
        };
        DataGrid.CellValuePushed += (_, e) =>
        {
            if (refreshing) return;
            Run(() => Document.SetCells(Sheet, new[] { (e.RowIndex + 1, e.ColumnIndex, e.Value?.ToString() ?? "") }));
        };
        DataGrid.CellFormatting += (_, e) =>
        {
            var cell = Sheet.Cell(e.RowIndex + 1, e.ColumnIndex);
            if (cell.IsFormula)
            {
                e.CellStyle!.ForeColor = StudioTheme.Accent;
                e.CellStyle.BackColor = StudioTheme.FormulaBg;
                if (!cell.HasCachedValue) { e.Value = "无缓存结果"; e.FormattingApplied = true; }
            }
            else if (!TypeInference.IsValid(cell.Text, Sheet.Columns[e.ColumnIndex].ValueType ?? "string", out string _))
            {
                e.CellStyle!.BackColor = StudioTheme.DangerBg;
                e.CellStyle.ForeColor = StudioTheme.DangerText;
            }
            if (Document.ReadOnlyReason != null || Sheet.ProtectionAt(e.RowIndex + 1, e.ColumnIndex) != null)
            {
                e.CellStyle!.Padding = new Padding(10, 0, 26, 0);
                e.CellStyle.BackColor = StudioTheme.FormulaBg;
            }
        };
        DataGrid.CellPainting += (_, e) =>
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0 || Document.ReadOnlyReason == null && Sheet.ProtectionAt(e.RowIndex + 1, e.ColumnIndex) == null) return;
            e.Paint(e.ClipBounds, DataGridViewPaintParts.All);
            e.Graphics!.SmoothingMode = SmoothingMode.AntiAlias;
            var circle = new Rectangle(e.CellBounds.Right - 20, e.CellBounds.Top + (e.CellBounds.Height - 14) / 2, 14, 14);
            using var pen = new Pen(StudioTheme.SubtleText, 1.1f);
            e.Graphics.DrawEllipse(pen, circle);
            using var questionFont = new Font(Font.FontFamily, 7.5F);
            TextRenderer.DrawText(e.Graphics, "?", questionFont, circle, StudioTheme.Muted, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
            e.Handled = true;
        };
        DataGrid.CellMouseClick += (_, e) =>
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && e.X >= DataGrid.Columns[e.ColumnIndex].Width - 24 &&
                (Document.ReadOnlyReason != null || Sheet.ProtectionAt(e.RowIndex + 1, e.ColumnIndex) != null)) ShowProtection();
        };
        DataGrid.CellToolTipTextNeeded += (_, e) =>
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            var cell = Sheet.Cell(e.RowIndex + 1, e.ColumnIndex);
            e.ToolTipText = cell.IsFormula ? cell.Text + "\n显示缓存结果，数据修改后需在 Excel 中重算。" : Sheet.ProtectionAt(e.RowIndex + 1, e.ColumnIndex) ?? cell.Text;
        };
        DataGrid.RowPostPaint += (_, e) => TextRenderer.DrawText(e.Graphics, (e.RowIndex + 2).ToString(), DataGrid.RowHeadersDefaultCellStyle.Font,
            new Rectangle(e.RowBounds.Left + 2, e.RowBounds.Top, DataGrid.RowHeadersWidth - 10, e.RowBounds.Height), StudioTheme.SubtleText, TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
        DataGrid.CurrentCellChanged += (_, _) => { if (refreshing) return; UpdateFormulaBar(); ColumnSelected?.Invoke(this, EventArgs.Empty); };
        DataGrid.ColumnHeaderMouseClick += (_, e) =>
        {
            if (e.ColumnIndex < 0) return;
            headerSelection = e.ColumnIndex;
            if (DataGrid.RowCount > 0) DataGrid.CurrentCell = DataGrid[e.ColumnIndex, 0];
            ColumnSelected?.Invoke(this, EventArgs.Empty);
        };
        DataGrid.KeyDown += (_, e) =>
        {
            if (DataGrid.IsCurrentCellInEditMode) return;
            if (e.Control && e.KeyCode == Keys.V) { e.SuppressKeyPress = true; Paste(); }
            else if (e.KeyCode == Keys.Delete) { e.SuppressKeyPress = true; ClearSelection(); }
        };
        DataGrid.DataError += (_, e) => { e.ThrowException = false; };
    }

    private void ConfigureStructureGrid()
    {
        structure.Dock = DockStyle.Fill; StudioTheme.Grid(structure); structure.VirtualMode = true; structure.RowHeadersVisible = false;
        structure.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "字段名", Width = 190, ReadOnly = Document.ReadOnlyReason != null || Sheet.HasTable });
        structure.Columns.Add(new DataGridViewComboBoxColumn { HeaderText = "数据类型", Width = 125, DataSource = UiDefaults.AllowedTypes });
        structure.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "主键序列 (-1 为无)", Width = 155 });
        structure.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "自定义属性", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 160, ReadOnly = true });
        structure.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "配置状态", Width = 200, ReadOnly = true });
        structure.CellValueNeeded += (_, e) =>
        {
            if (e.RowIndex >= Sheet.Columns.Count) return;
            var c = Sheet.Columns[e.RowIndex];
            e.Value = e.ColumnIndex switch { 0 => c.ColName, 1 => c.ValueType, 2 => c.PkOrder, 3 => string.Join(" · ", (c.ColProperties ?? new()).Select(p => p.PropertyName + "=" + p.PropertyValue)), _ => ColumnConfiguration.Missing(c, Registry) ?? "已配置" };
        };
        structure.CellValidating += (_, e) =>
        {
            if (refreshing || e.ColumnIndex != 2) return;
            if (!int.TryParse(e.FormattedValue?.ToString(), out int order) || order < -1)
            { e.Cancel = true; structure[e.ColumnIndex, e.RowIndex].ErrorText = "请输入 -1 或非负整数。"; }
            else structure[e.ColumnIndex, e.RowIndex].ErrorText = "";
        };
        structure.CellValuePushed += (_, e) =>
        {
            if (refreshing) return;
            var c = Sheet.Columns[e.RowIndex]; string value = e.Value?.ToString() ?? "";
            Run(() =>
            {
                if (e.ColumnIndex == 0) Document.RenameColumn(Sheet, e.RowIndex, value);
                else if (e.ColumnIndex == 1 && (c.ValueType != value || !c.IsConfigured))
                {
                    string? before = c.ValueType; bool configured = c.IsConfigured;
                    Document.Execute("修改类型", () => { c.ValueType = value; c.IsConfigured = true; }, () => { c.ValueType = before; c.IsConfigured = configured; });
                }
                else if (e.ColumnIndex == 2 && int.TryParse(value, out int order) && c.PkOrder != order)
                { int before = c.PkOrder; Document.Execute("修改主键", () => c.PkOrder = order, () => c.PkOrder = before); }
            });
        };
        structure.CurrentCellChanged += (_, _) => { if (!refreshing) ColumnSelected?.Invoke(this, EventArgs.Empty); };
        structure.DataError += (_, e) => e.ThrowException = false;
    }

    public bool CommitEdit() => DataGrid.EndEdit() && structure.EndEdit();

    public void RefreshModel()
    {
        if (IsDisposed || refreshing) return;
        refreshing = true;
        try
        {
            int selectedCol = DataGrid.CurrentCell?.ColumnIndex ?? 0, selectedRow = DataGrid.CurrentCell?.RowIndex ?? 0;
            if (DataGrid.ColumnCount != Sheet.Columns.Count)
            {
                DataGrid.Columns.Clear();
                foreach (var c in Sheet.Columns) DataGrid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = c.ColName, Width = 160, SortMode = DataGridViewColumnSortMode.NotSortable });
            }
            int unconfigured = 0;
            for (int i = 0; i < Sheet.Columns.Count; i++)
            {
                var missing = ColumnConfiguration.Missing(Sheet.Columns[i], Registry);
                if (missing != null) unconfigured++;
                string header = (missing == null ? "" : "● ") + Sheet.Columns[i].ColName + "  " + Sheet.Columns[i].ValueType;
                DataGrid.Columns[i].HeaderText = header;
                DataGrid.Columns[i].ToolTipText = missing ?? "字段类型已配置";
                DataGrid.Columns[i].HeaderCell.Style.ForeColor = missing == null ? StudioTheme.Text : StudioTheme.WarningText;
                int desiredW = TextRenderer.MeasureText(header, DataGrid.ColumnHeadersDefaultCellStyle.Font).Width + 32;
                if (DataGrid.Columns[i].Width < desiredW) DataGrid.Columns[i].Width = Math.Min(260, desiredW);
            }
            breadcrumbMeta.Text = $"{Sheet.Columns.Count} 字段 · {Sheet.RowCount:N0} 行" + (unconfigured > 0 ? $" · {unconfigured} 待配置" : "");
            breadcrumbMeta.ForeColor = unconfigured > 0 ? StudioTheme.WarningText : StudioTheme.Muted;

            bool editable = Document.ReadOnlyReason == null && !Sheet.IsProtected;
            addRowBtn.Enabled = editable;
            addColBtn.Enabled = editable && !Sheet.HasTable;

            DataGrid.RowCount = Sheet.RowCount;
            structure.RowCount = Sheet.Columns.Count;
            structure.ReadOnly = Sheet.SchemaIssue != null;
            var types = (DataGridViewComboBoxColumn)structure.Columns[1];
            var allowed = UiDefaults.AllowedTypes.Concat(Sheet.Columns.Select(c => c.ValueType ?? "string")).Distinct().ToArray();
            if (types.DataSource is not string[] existing || !existing.SequenceEqual(allowed)) types.DataSource = allowed;
            if (DataGrid.RowCount > 0 && DataGrid.ColumnCount > 0 && DataGrid.CurrentCell == null)
                DataGrid.CurrentCell = DataGrid[Math.Min(selectedCol, DataGrid.ColumnCount - 1), Math.Min(selectedRow, DataGrid.RowCount - 1)];
            DataGrid.Invalidate(); structure.Invalidate();
            if (viewSwitch.SelectedIndex == 2) preview.Text = DocumentStore.SchemaJson(Document);
            notice.Text = Sheet.SchemaIssue ?? Document.ReadOnlyReason ?? (Document.IsCsv ? $"CSV · {Document.Encoding.WebName} · 普通文本编辑，= 开头的内容不会作为公式执行。" :
                Document.HasExternalLinks ? "含外部链接 · 公式和引用已锁定 · 显示上次保存结果，数据修改后需在 Excel / WPS 中重算。" : "普通单元格可编辑 · 公式及特殊区域已锁定 · 显示上次保存结果，数据修改后需在 Excel / WPS 中重算");
            UpdateFormulaBar();
        }
        finally { refreshing = false; }
    }

    private void DocumentChanged(object? sender, EventArgs e)
    {
        if (IsHandleCreated && !IsDisposed) BeginInvoke((Action)RefreshModel);
    }

    private void UpdateFormulaBar()
    {
        if (DataGrid.CurrentCell is not { } current)
        {
            formula.Text = ""; formula.ReadOnly = true; ProtectionMessage = null; protectionHelp.Visible = false; lockState.Text = "";
            return;
        }
        var cell = Sheet.Cell(current.RowIndex + 1, current.ColumnIndex);
        address.Text = EditorSheet.Address(current.RowIndex + 1, current.ColumnIndex);
        formula.Text = cell.Text;
        ProtectionMessage = Document.ReadOnlyReason ?? Sheet.ProtectionAt(current.RowIndex + 1, current.ColumnIndex);
        formula.ReadOnly = true;
        lockState.Text = ProtectionMessage != null ? "已锁定" : "可编辑";
        lockState.ForeColor = ProtectionMessage != null ? StudioTheme.WarningText : StudioTheme.Muted;
        protectionHelp.Visible = ProtectionMessage != null;
        formula.ForeColor = cell.IsFormula ? StudioTheme.Accent : StudioTheme.Text;
    }

    private void ShowProtection()
    {
        if (ProtectionMessage == null) return;
        var current = DataGrid.CurrentCell;
        var cell = current == null ? null : Sheet.Cell(current.RowIndex + 1, current.ColumnIndex);
        MessageBox.Show(this, ProtectionMessage + (cell?.IsFormula == true ? "\n\n公式：" + cell.Text : "") +
            "\n\n请在 Excel / WPS 中修改、更新链接并保存，再重新加载。文件里的缓存结果可能不是最新值。", "受保护数据", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }

    public void Paste()
    {
        if (DataGrid.CurrentCell == null || !Clipboard.ContainsText()) return;
        Run(() =>
        {
            using var reader = new CsvReader(new StringReader(Clipboard.GetText()), new CsvConfiguration(CultureInfo.InvariantCulture) { Delimiter = "\t", HasHeaderRecord = false, IgnoreBlankLines = false });
            var edits = new List<(int Row, int Column, string Text)>(); int offset = 0;
            while (reader.Read())
            {
                for (int col = 0; col < reader.Parser.Count; col++) edits.Add((SelectedDataRow + offset, DataGrid.CurrentCell.ColumnIndex + col, reader.GetField(col) ?? ""));
                offset++;
            }
            if (edits.Count > 0) Document.SetCells(Sheet, edits);
        });
    }

    public void ClearSelection() => Run(() => Document.SetCells(Sheet,
        DataGrid.SelectedCells.Cast<DataGridViewCell>().Select(c => (c.RowIndex + 1, c.ColumnIndex, "")).ToList()));

    private void Run(Action action)
    {
        try { action(); }
        catch (Exception error) { MessageBox.Show(this, error.Message, "无法编辑", MessageBoxButtons.OK, MessageBoxIcon.Information); }
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) Document.Changed -= DocumentChanged;
        base.Dispose(disposing);
    }
}
