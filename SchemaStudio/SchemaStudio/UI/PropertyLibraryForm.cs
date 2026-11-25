using System.Drawing.Drawing2D;
using System.Text.Json;
using SchemaStudio.Models;
using SchemaStudio.Services;

namespace SchemaStudio.UI
{
    public class PropertyLibraryForm : Form
    {
        private readonly ListBox list;
        private readonly TextBox txtKey, txtTitle, txtEnum, txtDefault;
        private readonly ComboBox cmbKind, cmbScope;
        private readonly Label lblEnum, lblEnumHint, lblEmpty;
        private readonly TableLayoutPanel editorGrid;
        private readonly StudioButton btnDelete;
        private readonly PropertyRegistry working;
        private bool loading;

        public PropertyRegistry Result => working;

        public PropertyLibraryForm(PropertyRegistry reg)
        {
            Text = "自定义属性库";
            ClientSize = new Size(780, 500);
            MinimumSize = new Size(640, 420);
            StartPosition = FormStartPosition.CenterParent;
            Font = StudioTheme.UiFont(9F);
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = StudioTheme.Surface;
            ForeColor = StudioTheme.Text;
            HandleCreated += (_, _) => StudioTheme.ApplyModernTitleBar(this);

            working = JsonSerializer.Deserialize<PropertyRegistry>(
                          JsonSerializer.Serialize(reg, DataIO.JsonOptions),
                          DataIO.JsonOptions) ?? new PropertyRegistry();

            var body = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Padding = new Padding(14) };
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 260));
            body.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            // ── Left: List + Direct "+ 新建属性" / "删除" toolbar ──
            var leftCard = new StudioCard { Dock = DockStyle.Fill, BackColor = StudioTheme.Background, Padding = new Padding(1) };

            var leftInner = new Panel { Dock = DockStyle.Fill, BackColor = StudioTheme.Background };
            var leftBar = new StudioCard { Dock = DockStyle.Top, BorderAll = false, BorderBottom = true, BackColor = StudioTheme.Background, Padding = new Padding(8, 5, 6, 6) };

            var btnNew = StudioTheme.Button("新建属性", (_, _) => AddNewProperty(), icon: StudioIcon.TableAdd, borderless: true);
            btnNew.Dock = DockStyle.Left;
            btnDelete = StudioTheme.IconButton(StudioIcon.Trash, "删除选中属性", (_, _) => DeleteSelected());
            btnDelete.Dock = DockStyle.Right;
            leftBar.Height = Math.Max(btnNew.GetPreferredSize(Size.Empty).Height, btnDelete.GetPreferredSize(Size.Empty).Height) + 12;
            leftBar.Controls.Add(btnNew);
            leftBar.Controls.Add(btnDelete);

            list = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                IntegralHeight = false,
                DrawMode = DrawMode.OwnerDrawFixed,
                BackColor = StudioTheme.Background,
                ForeColor = StudioTheme.Text
            };
            UpdateListMetrics();
            list.DrawItem += DrawListItem;
            list.SelectedIndexChanged += (_, _) => LoadSelected();
            Load += (_, _) => { UpdateListMetrics(); list.Invalidate(); };
            DpiChanged += (_, _) => { UpdateListMetrics(); list.Invalidate(); };

            leftInner.Controls.Add(list);
            leftInner.Controls.Add(leftBar);
            leftCard.Controls.Add(leftInner);
            body.Controls.Add(leftCard, 0, 0);

            // ── Right: Live detail editor (auto-syncs to selected item) ──
            var rightCard = new StudioCard { Dock = DockStyle.Fill, BackColor = StudioTheme.Surface, Padding = new Padding(20, 14, 20, 14), Margin = new Padding(12, 0, 0, 0) };

            editorGrid = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 2 };
            editorGrid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            editorGrid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            Label AddRow(string label, Control input)
            {
                int r = editorGrid.RowCount++;
                editorGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                var lbl = new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = StudioTheme.Muted, Margin = new Padding(0, 9, 14, 9) };
                input.Dock = DockStyle.Fill;
                input.Margin = new Padding(0, 5, 0, 5);
                editorGrid.Controls.Add(lbl, 0, r);
                editorGrid.Controls.Add(input, 1, r);
                return lbl;
            }

            txtKey = new TextBox { BorderStyle = BorderStyle.FixedSingle };
            txtTitle = new TextBox { BorderStyle = BorderStyle.FixedSingle };
            cmbKind = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
            cmbKind.Items.AddRange(new[] { "enum", "bool", "int", "string" });
            cmbScope = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
            cmbScope.Items.AddRange(new[] { "column", "table", "both" });
            txtEnum = new TextBox { BorderStyle = BorderStyle.FixedSingle, PlaceholderText = "例如：ALL, C, S" };
            txtDefault = new TextBox { BorderStyle = BorderStyle.FixedSingle };

            AddRow("属性键 (key)", txtKey);
            AddRow("显示标题", txtTitle);
            AddRow("属性类型", cmbKind);
            AddRow("适用范围", cmbScope);
            lblEnum = AddRow("枚举值", txtEnum);
            AddRow("默认值", txtDefault);

            lblEnumHint = new Label
            {
                Text = "提示：修改左侧选中项会实时生效；scope 选 table 或 both 才会出现在「表属性」中。",
                AutoSize = true,
                MaximumSize = new Size(420, 0),
                ForeColor = StudioTheme.SubtleText,
                Font = StudioTheme.UiFont(8.25F),
                Margin = new Padding(0, 14, 0, 0)
            };
            int hintRow = editorGrid.RowCount++;
            editorGrid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            editorGrid.Controls.Add(lblEnumHint, 0, hintRow);
            editorGrid.SetColumnSpan(lblEnumHint, 2);

            txtKey.TextChanged += (_, _) => SyncCurrentFromInputs();
            txtTitle.TextChanged += (_, _) => SyncCurrentFromInputs();
            cmbKind.SelectedIndexChanged += (_, _) => { UpdateEnumVisibility(); SyncCurrentFromInputs(); };
            cmbScope.SelectedIndexChanged += (_, _) => SyncCurrentFromInputs();
            txtEnum.TextChanged += (_, _) => SyncCurrentFromInputs();
            txtDefault.TextChanged += (_, _) => SyncCurrentFromInputs();

            lblEmpty = new Label
            {
                Text = "暂无自定义属性，点击左侧「新建属性」开始添加。",
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = StudioTheme.Muted,
                Visible = false
            };

            rightCard.Controls.Add(editorGrid);
            rightCard.Controls.Add(lblEmpty);
            body.Controls.Add(rightCard, 1, 0);

            // ── Footer ──
            var footer = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(14, 10, 14, 12), BackColor = StudioTheme.Background };
            footer.Resize += (_, _) => footer.Invalidate();
            footer.Paint += (_, e) => { using var p = new Pen(StudioTheme.Border); e.Graphics.DrawLine(p, 0, 0, footer.Width, 0); };
            var btnOk = StudioTheme.Button("保存属性库", (_, _) => ConfirmAndClose(), primary: true, icon: StudioIcon.Save);
            var btnCancel = StudioTheme.Button("取消", (_, _) => { DialogResult = DialogResult.Cancel; Close(); });
            footer.Controls.Add(btnOk);
            footer.Controls.Add(btnCancel);

            Controls.Add(body);
            Controls.Add(footer);
            AcceptButton = btnOk;
            CancelButton = btnCancel;

            RefreshList(working.Properties.FirstOrDefault());
        }

        private int titleH = 20, subH = 18;

        private void UpdateListMetrics()
        {
            using var titleFont = StudioTheme.UiFont(9F, FontStyle.Bold);
            using var subFont = StudioTheme.UiFont(8.25F);
            titleH = StudioTheme.LineHeight(titleFont);
            subH = StudioTheme.LineHeight(subFont);
            list.ItemHeight = titleH + subH + 18;
        }

        private void DrawListItem(object? sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= list.Items.Count) return;
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            using (var bg = new SolidBrush(StudioTheme.Background)) g.FillRectangle(bg, e.Bounds);

            bool selected = (e.State & DrawItemState.Selected) != 0;
            if (selected)
            {
                var pill = new Rectangle(e.Bounds.X + 6, e.Bounds.Y + 3, e.Bounds.Width - 12, e.Bounds.Height - 6);
                using var path = StudioTheme.RoundedRect(pill, 6);
                using var fill = new SolidBrush(StudioTheme.ActiveTab);
                g.FillPath(fill, path);
            }

            if (list.Items[e.Index] is not PropertyDef d) return;
            string primary = string.IsNullOrWhiteSpace(d.Title) ? (string.IsNullOrWhiteSpace(d.Key) ? "(未命名属性)" : d.Key) : d.Title;
            string scopeLabel = d.Scope switch { "table" => "仅表", "both" => "表/列", _ => "仅列" };
            string secondary = $"{(string.IsNullOrWhiteSpace(d.Key) ? "未填 key" : d.Key)}  ·  {d.Kind}  ·  {scopeLabel}";

            using var titleFont = StudioTheme.UiFont(9F, selected ? FontStyle.Bold : FontStyle.Regular);
            using var subFont = StudioTheme.UiFont(8.25F);
            int top = e.Bounds.Y + Math.Max(4, (e.Bounds.Height - (titleH + subH + 2)) / 2);
            TextRenderer.DrawText(g, primary, titleFont,
                new Rectangle(e.Bounds.X + 14, top, e.Bounds.Width - 28, titleH),
                StudioTheme.Text, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            TextRenderer.DrawText(g, secondary, subFont,
                new Rectangle(e.Bounds.X + 14, top + titleH + 2, e.Bounds.Width - 28, subH),
                StudioTheme.Muted, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
        }

        private void AddNewProperty()
        {
            int n = 1;
            string key;
            do { key = $"prop_{n++}"; }
            while (working.Properties.Any(p => p.Key.Equals(key, StringComparison.OrdinalIgnoreCase)));

            var item = new PropertyDef
            {
                Key = key,
                Title = "新属性",
                Kind = "string",
                Scope = "both",
                EnumValues = new List<string>(),
                Default = ""
            };
            working.Properties.Add(item);
            RefreshList(item);
            txtKey.Focus();
            txtKey.SelectAll();
        }

        private void DeleteSelected()
        {
            if (list.SelectedItem is not PropertyDef d) return;
            int idx = working.Properties.IndexOf(d);
            working.Properties.Remove(d);
            PropertyDef? next = working.Properties.Count == 0 ? null : working.Properties[Math.Min(idx, working.Properties.Count - 1)];
            RefreshList(next);
        }

        private void RefreshList(PropertyDef? select)
        {
            loading = true;
            try
            {
                list.BeginUpdate();
                list.Items.Clear();
                foreach (var p in working.Properties) list.Items.Add(p);
                list.EndUpdate();
                if (select != null && list.Items.Contains(select)) list.SelectedItem = select;
                else if (list.Items.Count > 0) list.SelectedIndex = 0;
            }
            finally { loading = false; }
            LoadSelected();
        }

        private void LoadSelected()
        {
            loading = true;
            try
            {
                if (list.SelectedItem is not PropertyDef d)
                {
                    editorGrid.Visible = false;
                    lblEmpty.Visible = true;
                    btnDelete.Enabled = false;
                    return;
                }
                editorGrid.Visible = true;
                lblEmpty.Visible = false;
                btnDelete.Enabled = true;

                txtKey.Text = d.Key;
                txtTitle.Text = d.Title;
                cmbKind.SelectedItem = cmbKind.Items.Contains(d.Kind) ? d.Kind : "string";
                cmbScope.SelectedItem = cmbScope.Items.Contains(d.Scope) ? d.Scope : "column";
                txtEnum.Text = string.Join(", ", d.EnumValues ?? new List<string>());
                txtDefault.Text = d.Default ?? "";
                UpdateEnumVisibility();
            }
            finally { loading = false; }
        }

        private void UpdateEnumVisibility()
        {
            bool isEnum = string.Equals(cmbKind.SelectedItem?.ToString(), "enum", StringComparison.OrdinalIgnoreCase);
            lblEnum.Visible = isEnum;
            txtEnum.Visible = isEnum;
        }

        private void SyncCurrentFromInputs()
        {
            if (loading || list.SelectedItem is not PropertyDef d) return;
            d.Key = txtKey.Text.Trim();
            d.Title = txtTitle.Text.Trim();
            d.Kind = cmbKind.SelectedItem?.ToString() ?? "string";
            d.Scope = cmbScope.SelectedItem?.ToString() ?? "column";
            d.EnumValues = d.Kind == "enum"
                ? txtEnum.Text.Split(new[] { ',', '，' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).Distinct().ToList()
                : new List<string>();
            d.Default = txtDefault.Text.Trim();
            list.Invalidate();
        }

        private void ConfirmAndClose()
        {
            SyncCurrentFromInputs();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in working.Properties)
            {
                if (string.IsNullOrWhiteSpace(p.Key))
                {
                    list.SelectedItem = p;
                    MessageBox.Show(this, "存在属性键 (key) 为空的条目，请填写或删除后再保存。", "属性库校验", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtKey.Focus();
                    return;
                }
                if (!seen.Add(p.Key))
                {
                    list.SelectedItem = p;
                    MessageBox.Show(this, $"属性键 \"{p.Key}\" 重复，请修改为唯一的 key。", "属性库校验", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtKey.Focus();
                    return;
                }
                if (p.Kind == "enum" && (p.EnumValues == null || p.EnumValues.Count == 0))
                {
                    list.SelectedItem = p;
                    MessageBox.Show(this, $"枚举属性 \"{p.Key}\" 请至少填写一个枚举值（逗号分隔）。", "属性库校验", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    txtEnum.Focus();
                    return;
                }
            }
            DialogResult = DialogResult.OK;
            Close();
        }
    }
}
