using System.ComponentModel;
using SchemaStudio.Models;

namespace SchemaStudio.UI
{
    public class NewTableForm : Form
    {
        private readonly ComboBox cmbDir, cmbFormat;
        private readonly TextBox txtName;
        private readonly DataGridView gridCols, gridSheets;
        private readonly Panel sheetsSection;
        private readonly Label lblTablePropSummary;

        private readonly BindingList<SheetNameRow> sheets = new();
        private readonly Dictionary<string, BindingList<NewColRow>> sheetCols = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, List<PropertyKV>> sheetTableProps = new(StringComparer.OrdinalIgnoreCase);

        private string CurrentSheet => GetSelectedSheetNameOrFirst();
        private readonly PropertyRegistry registry;
        private readonly string root;
        private readonly string[] allowed;
        private string? _oldSheetNameEditing;

        public string? ResultDir { get; private set; }
        public string? ResultFileName { get; private set; }
        public string ResultFormat { get; private set; } = "csv";
        public List<string> ResultSheets => sheets.Select(s => s.Name).Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
        public Dictionary<string, List<NewColRow>> ResultSheetColumns =>
            sheetCols.ToDictionary(kv => kv.Key, kv => kv.Value.ToList(), StringComparer.OrdinalIgnoreCase);
        public IEnumerable<NewColRow> Columns =>
            sheetCols.TryGetValue(CurrentSheet, out var list) ? list : Enumerable.Empty<NewColRow>();
        public Dictionary<string, List<PropertyKV>> ResultSheetTableProps =>
            sheetTableProps.ToDictionary(kv => kv.Key, kv => kv.Value?.ToList() ?? new List<PropertyKV>(), StringComparer.OrdinalIgnoreCase);

        public NewTableForm(string root, List<string> allDirs, PropertyRegistry reg, string[] allowedTypes)
        {
            this.root = root; registry = reg; allowed = allowedTypes;

            Text = "新建表";
            ClientSize = new Size(960, 640);
            MinimumSize = new Size(760, 520);
            StartPosition = FormStartPosition.CenterParent;
            Font = StudioTheme.UiFont(9F);
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = StudioTheme.Surface;
            ForeColor = StudioTheme.Text;
            HandleCreated += (_, _) => StudioTheme.ApplyModernTitleBar(this);

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4, Padding = new Padding(16, 12, 16, 10) };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            // ── Top Meta Bar ──
            var meta = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 5, Margin = new Padding(0, 0, 0, 10) };
            meta.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            meta.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            meta.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            meta.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            meta.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            cmbDir = new ComboBox { Dock = DockStyle.Fill, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbDir.Items.AddRange(allDirs
                .Select(d => d.Replace(root, "").TrimStart(Path.DirectorySeparatorChar))
                .Select(s => string.IsNullOrEmpty(s) ? "(根目录)" : s)
                .ToArray<object>());
            cmbDir.SelectedIndex = 0;

            var btnBrowse = StudioTheme.Button("浏览...", (_, _) =>
            {
                using var fbd = new FolderBrowserDialog { SelectedPath = root };
                if (fbd.ShowDialog(this) == DialogResult.OK)
                {
                    var p = fbd.SelectedPath;
                    var show = p.StartsWith(root, StringComparison.OrdinalIgnoreCase)
                        ? p.Replace(root, "").TrimStart(Path.DirectorySeparatorChar)
                        : p;
                    if (string.IsNullOrEmpty(show)) show = "(根目录)";
                    if (!cmbDir.Items.Cast<object>().Any(x => string.Equals(x.ToString(), show, StringComparison.OrdinalIgnoreCase)))
                        cmbDir.Items.Add(show);
                    cmbDir.SelectedItem = show;
                }
            });

            txtName = new TextBox { Dock = DockStyle.Fill, Text = "new_table", BorderStyle = BorderStyle.FixedSingle };
            cmbFormat = new ComboBox { Width = 110, DropDownStyle = ComboBoxStyle.DropDownList };
            cmbFormat.Items.AddRange(new[] { "csv", "xlsx", "xls" });
            cmbFormat.SelectedIndex = 0;

            meta.Controls.Add(new Label { Text = "保存目录", AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = StudioTheme.Muted, Margin = new Padding(0, 6, 8, 6) }, 0, 0);
            meta.Controls.Add(cmbDir, 1, 0); meta.SetColumnSpan(cmbDir, 3);
            meta.Controls.Add(btnBrowse, 4, 0);

            meta.Controls.Add(new Label { Text = "文件名", AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = StudioTheme.Muted, Margin = new Padding(0, 8, 8, 4) }, 0, 1);
            meta.Controls.Add(txtName, 1, 1);
            meta.Controls.Add(new Label { Text = "格式", AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = StudioTheme.Muted, Margin = new Padding(14, 8, 8, 4) }, 2, 1);
            meta.Controls.Add(cmbFormat, 3, 1);
            layout.Controls.Add(meta, 0, 0);

            // ── Columns Grid ──
            gridCols = new DataGridView { Dock = DockStyle.Fill };
            StudioTheme.Grid(gridCols);
            gridCols.RowHeadersVisible = false;
            gridCols.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            gridCols.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "列名", DataPropertyName = nameof(NewColRow.ColName), Width = 220 });
            gridCols.Columns.Add(new DataGridViewComboBoxColumn { HeaderText = "数据类型", DataPropertyName = nameof(NewColRow.ValueType), Width = 130, DataSource = allowed });
            gridCols.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "主键序列", DataPropertyName = nameof(NewColRow.PkOrderText), Width = 110 });
            gridCols.Columns.Add(new DataGridViewButtonColumn { HeaderText = "列属性", Text = "属性…", UseColumnTextForButtonValue = true, Width = 84 });
            gridCols.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "属性摘要", DataPropertyName = nameof(NewColRow.PropsSummary), AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, ReadOnly = true });
            gridCols.CellContentClick += (_, e) =>
            {
                if (e.RowIndex < 0) return;
                if (gridCols.Columns[e.ColumnIndex] is DataGridViewButtonColumn &&
                    gridCols.Rows[e.RowIndex].DataBoundItem is NewColRow row)
                {
                    var defs = registry.Properties.Where(d => d.Scope == "column" || d.Scope == "both").ToList();
                    using var dlg = new AssignPropertyForm(defs, row.ColProperties ?? new List<PropertyKV>());
                    if (dlg.ShowDialog(this) == DialogResult.OK)
                    {
                        row.ColProperties = dlg.Result;
                        row.PropsSummary = string.Join("; ", row.ColProperties.Select(p => $"{p.PropertyName}={p.PropertyValue}"));
                        gridCols.InvalidateRow(e.RowIndex);
                    }
                }
            };
            var gridWrap = new Panel { Dock = DockStyle.Fill, Padding = new Padding(1), BackColor = StudioTheme.Border };
            gridWrap.Controls.Add(gridCols);
            layout.Controls.Add(gridWrap, 0, 1);

            // ── Column Actions Bar ──
            var colActions = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, WrapContents = true, Padding = new Padding(0, 10, 0, 4) };
            colActions.Controls.Add(StudioTheme.Button("添加列", (_, _) =>
            {
                var list = GetCurrentCols();
                list.Add(new NewColRow { ColName = $"col_{list.Count + 1}", ValueType = "string", PkOrder = -1 });
            }, icon: StudioIcon.ColAdd));
            colActions.Controls.Add(StudioTheme.Button("删除选中列", (_, _) =>
            {
                var list = GetCurrentCols();
                foreach (DataGridViewRow r in gridCols.SelectedRows)
                    if (r.DataBoundItem is NewColRow x) list.Remove(x);
            }, icon: StudioIcon.Trash));
            colActions.Controls.Add(StudioTheme.Button("表属性…", (_, _) =>
            {
                var defs = registry.Properties.Where(d => d.Scope == "table" || d.Scope == "both").ToList();
                var key = CurrentSheet;
                if (!sheetTableProps.TryGetValue(key, out var list)) sheetTableProps[key] = list = new List<PropertyKV>();
                using var dlg = new AssignPropertyForm(defs, list) { Text = "表属性" };
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    sheetTableProps[key] = dlg.Result ?? new List<PropertyKV>();
                    RefreshTablePropSummaryFor(key);
                }
            }, icon: StudioIcon.Sliders));
            lblTablePropSummary = new Label { Text = "表属性：—", AutoSize = true, ForeColor = StudioTheme.Muted, Padding = new Padding(8, 7, 0, 0) };
            colActions.Controls.Add(lblTablePropSummary);
            layout.Controls.Add(colActions, 0, 2);

            // ── Excel Sheets Section ──
            sheetsSection = new Panel { Dock = DockStyle.Top, Height = 112, Visible = false, Padding = new Padding(0, 6, 0, 0) };
            var sheetLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2 };
            sheetLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            sheetLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            gridSheets = new DataGridView { Dock = DockStyle.Fill, DataSource = sheets };
            StudioTheme.Grid(gridSheets);
            gridSheets.RowHeadersVisible = false;
            gridSheets.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            gridSheets.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Sheet 名", DataPropertyName = nameof(SheetNameRow.Name), AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            gridSheets.CellBeginEdit += (_, e) => { if (e.RowIndex >= 0 && gridSheets.Rows[e.RowIndex].DataBoundItem is SheetNameRow s) _oldSheetNameEditing = s.Name; };
            gridSheets.SelectionChanged += (_, _) => { if (!gridSheets.IsCurrentCellInEditMode) { BindColsToCurrentSheet(); RefreshTablePropSummaryFor(CurrentSheet); } };
            gridSheets.CellEndEdit += (_, e) =>
            {
                if (e.RowIndex < 0) return;
                var s = (SheetNameRow)gridSheets.Rows[e.RowIndex].DataBoundItem!;
                var newName = (s.Name ?? "").Trim();
                if (string.IsNullOrWhiteSpace(newName))
                {
                    MessageBox.Show("Sheet 名不能为空");
                    s.Name = gridSheets.Rows[e.RowIndex].Tag as string ?? "Sheet1";
                    gridSheets.InvalidateRow(e.RowIndex);
                    return;
                }
                var oldName = _oldSheetNameEditing ?? (gridSheets.Rows[e.RowIndex].Tag as string ?? newName);
                _oldSheetNameEditing = null;
                if (!string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase) && sheetCols.ContainsKey(newName))
                {
                    MessageBox.Show("Sheet 名重复");
                    s.Name = oldName;
                    gridSheets.InvalidateRow(e.RowIndex);
                    return;
                }
                if (!sheetCols.TryGetValue(oldName, out var colsList)) colsList = new BindingList<NewColRow>();
                sheetCols.Remove(oldName); sheetCols[newName] = colsList;
                if (sheetTableProps.Remove(oldName, out var tp)) sheetTableProps[newName] = tp;
                if (!sheetTableProps.ContainsKey(newName)) sheetTableProps[newName] = new List<PropertyKV>();
                gridSheets.Rows[e.RowIndex].Tag = newName;
                gridCols.DataSource = sheetCols[newName];
                RefreshTablePropSummaryFor(newName);
            };
            var sheetWrap = new Panel { Dock = DockStyle.Fill, Padding = new Padding(1), BackColor = StudioTheme.Border };
            sheetWrap.Controls.Add(gridSheets);

            var sheetBtns = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, FlowDirection = FlowDirection.TopDown, Padding = new Padding(10, 0, 0, 0) };
            sheetBtns.Controls.Add(StudioTheme.Button("添加 Sheet", (_, _) =>
            {
                var name = AutoSheetName();
                sheets.Add(new SheetNameRow { Name = name });
                sheetCols[name] = new BindingList<NewColRow>
                {
                    new() { ColName = "id", ValueType = "int", PkOrder = 1 },
                    new() { ColName = "desc", ValueType = "string", PkOrder = -1 }
                };
                gridSheets.ClearSelection();
                int last = gridSheets.Rows.Count - 1;
                if (last >= 0) { gridSheets.Rows[last].Selected = true; gridSheets.CurrentCell = gridSheets.Rows[last].Cells[0]; gridSheets.Rows[last].Tag = name; }
                if (!sheetTableProps.ContainsKey(name)) sheetTableProps[name] = new List<PropertyKV>();
                BindColsToCurrentSheet();
                RefreshTablePropSummaryFor(name);
            }));
            sheetBtns.Controls.Add(StudioTheme.Button("删除 Sheet", (_, _) =>
            {
                if (gridSheets.CurrentRow?.DataBoundItem is not SheetNameRow s) return;
                if (sheets.Count <= 1) { MessageBox.Show("至少保留一个 Sheet"); return; }
                sheetCols.Remove(s.Name); sheetTableProps.Remove(s.Name); sheets.Remove(s);
                BindColsToCurrentSheet();
                if (gridSheets.CurrentRow?.DataBoundItem is SheetNameRow cur) RefreshTablePropSummaryFor(cur.Name);
                else lblTablePropSummary.Text = "表属性：—";
            }));
            sheetLayout.Controls.Add(sheetWrap, 0, 0);
            sheetLayout.Controls.Add(sheetBtns, 1, 0);
            sheetsSection.Controls.Add(sheetLayout);
            layout.Controls.Add(sheetsSection, 0, 3);

            cmbFormat.SelectedIndexChanged += (_, _) =>
            {
                ResultFormat = cmbFormat.SelectedItem!.ToString()!;
                sheetsSection.Visible = ResultFormat is "xlsx" or "xls";
            };

            // ── Footer ──
            var footer = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(16, 10, 16, 12), BackColor = StudioTheme.Background };
            footer.Paint += (_, e) => { using var p = new Pen(StudioTheme.Border); e.Graphics.DrawLine(p, 0, 0, footer.Width, 0); };
            var btnOk = StudioTheme.Button("确定创建", (_, _) => OnOk(), primary: true);
            var btnCancel = StudioTheme.Button("取消", (_, _) => { DialogResult = DialogResult.Cancel; Close(); });
            footer.Controls.Add(btnOk);
            footer.Controls.Add(btnCancel);

            Controls.Add(layout);
            Controls.Add(footer);
            AcceptButton = btnOk;
            CancelButton = btnCancel;

            var first = "Sheet1";
            sheetTableProps[first] = new List<PropertyKV>();
            RefreshTablePropSummaryFor(first);
            sheets.Add(new SheetNameRow { Name = first });
            sheetCols[first] = new BindingList<NewColRow>
            {
                new() { ColName = "id", ValueType = "int", PkOrder = 1 },
                new() { ColName = "desc", ValueType = "string", PkOrder = -1 }
            };
            gridSheets.ClearSelection();
            if (gridSheets.Rows.Count > 0) { gridSheets.Rows[0].Selected = true; gridSheets.CurrentCell = gridSheets.Rows[0].Cells[0]; gridSheets.Rows[0].Tag = first; }
            BindColsToCurrentSheet();
        }

        private BindingList<NewColRow> GetCurrentCols()
        {
            var key = CurrentSheet;
            if (!sheetCols.TryGetValue(key, out var list)) { list = new BindingList<NewColRow>(); sheetCols[key] = list; }
            gridCols.DataSource = list;
            return list;
        }
        private void BindColsToCurrentSheet()
        {
            var key = CurrentSheet;
            if (!sheetCols.TryGetValue(key, out var list)) { list = new BindingList<NewColRow>(); sheetCols[key] = list; }
            gridCols.DataSource = list;
        }
        private string GetSelectedSheetNameOrFirst()
        {
            if (gridSheets.CurrentRow?.DataBoundItem is SheetNameRow cur && !string.IsNullOrWhiteSpace(cur.Name)) return cur.Name;
            if (gridSheets.Rows.Count > 0 && gridSheets.Rows[0].DataBoundItem is SheetNameRow first && !string.IsNullOrWhiteSpace(first.Name)) return first.Name;
            if (sheetCols.Count > 0) return sheetCols.Keys.First();
            return "Sheet1";
        }
        private string AutoSheetName()
        {
            int i = sheets.Count + 1; string name;
            do { name = $"Sheet{i++}"; } while (sheets.Any(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase)));
            return name;
        }
        private void OnOk()
        {
            if (!gridCols.EndEdit() || !gridSheets.EndEdit()) return;
            var name = (txtName.Text ?? "").Trim();
            if (string.IsNullOrEmpty(name)) { MessageBox.Show("文件名不能为空"); return; }
            if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) { MessageBox.Show("文件名包含非法字符"); return; }

            var rel = cmbDir.SelectedItem!.ToString()!;
            string destDir = rel == "(根目录)" ? root :
                (Path.IsPathRooted(rel) ? rel : Path.Combine(root, rel.TrimStart(Path.DirectorySeparatorChar)));
            if (!Directory.Exists(destDir)) { MessageBox.Show("目标目录无效"); return; }

            ResultFormat = cmbFormat.SelectedItem!.ToString()!;
            if (ResultFormat is "xlsx" or "xls")
            {
                if (sheets.Count == 0) { MessageBox.Show("请至少添加一个工作表"); return; }
                if (sheets.Any(s => string.IsNullOrWhiteSpace(s.Name))) { MessageBox.Show("存在空的工作表名"); return; }
                if (sheets.Select(s => s.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count() != sheets.Count) { MessageBox.Show("工作表名重复"); return; }
                foreach (var s in sheets)
                    if (!sheetCols.TryGetValue(s.Name, out var cols) || cols.Count == 0) { MessageBox.Show($"Sheet \"{s.Name}\" 至少需要一列"); return; }
            }
            else
            {
                if (!sheetCols.TryGetValue(CurrentSheet, out var cols) || cols.Count == 0) { MessageBox.Show("请至少添加一列"); return; }
            }

            var selectedSheets = ResultFormat == "csv" ? sheetCols.Where(p => p.Key == CurrentSheet) : sheetCols;
            foreach (var entry in selectedSheets)
            {
                if (entry.Value.Any(c => string.IsNullOrWhiteSpace(c.ColName)) || entry.Value.Select(c => c.ColName.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).Count() != entry.Value.Count)
                { MessageBox.Show(this, $"{entry.Key} 的字段名不能为空或重复。"); return; }
                foreach (var column in entry.Value)
                    if (string.IsNullOrWhiteSpace(column.ValueType)) { MessageBox.Show(this, $"请为字段 {column.ColName} 选择数据类型。"); return; }
            }
            ResultDir = destDir;
            ResultFileName = name;
            DialogResult = DialogResult.OK;
            Close();
        }
        private void RefreshTablePropSummaryFor(string sheet)
        {
            if (!sheetTableProps.TryGetValue(sheet, out var list) || list == null || list.Count == 0)
            {
                lblTablePropSummary.Text = "表属性：—";
                return;
            }
            lblTablePropSummary.Text = "表属性：" + string.Join("; ", list.Select(p => $"{p.PropertyName}={p.PropertyValue}"));
        }
    }
}
