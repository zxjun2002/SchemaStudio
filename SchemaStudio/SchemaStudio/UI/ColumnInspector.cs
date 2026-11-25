using System.Drawing.Drawing2D;
using SchemaStudio.Models;
using SchemaStudio.Services;

namespace SchemaStudio.UI;

internal sealed class ColumnInspector : UserControl
{
    private readonly Panel content = new() { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(16, 8, 16, 16), BackColor = StudioTheme.Surface };
    private readonly Label headerTitle = new() { Text = "字段属性", AutoSize = true, Font = StudioTheme.UiFont(9.5F, FontStyle.Bold), ForeColor = StudioTheme.Text };
    private readonly Label headerBadge = new() { AutoSize = true, Font = StudioTheme.UiFont(8.25F), ForeColor = StudioTheme.Muted, Padding = new Padding(6, 2, 0, 0) };
    private string fingerprint = "";
    private bool rendering;

    public ColumnInspector()
    {
        Dock = DockStyle.Fill;
        BackColor = StudioTheme.Surface;
        var header = new Panel { Dock = DockStyle.Top, Height = 38, BackColor = StudioTheme.Background, Padding = new Padding(14, 9, 12, 0) };
        header.Paint += (_, e) => { using var p = new Pen(StudioTheme.Border); e.Graphics.DrawLine(p, 0, header.Height - 1, header.Width, header.Height - 1); };
        var row = new FlowLayoutPanel { Dock = DockStyle.Fill, WrapContents = false, BackColor = Color.Transparent, Margin = Padding.Empty, Padding = Padding.Empty };
        row.Controls.Add(headerTitle);
        row.Controls.Add(headerBadge);
        header.Controls.Add(row);

        Controls.Add(content);
        Controls.Add(header);
    }

    public void ShowColumn(TableEditorControl? view, PropertyRegistry registry, bool force = false)
    {
        if (rendering) return;
        if (view?.Sheet.SchemaIssue is { } issue)
        {
            if (fingerprint == issue) return;
            fingerprint = issue; headerBadge.Text = ""; ClearContent();
            content.Controls.Add(new Label { Text = issue, Dock = DockStyle.Top, AutoSize = true, MaximumSize = new Size(240, 0), ForeColor = StudioTheme.Muted, Padding = new Padding(0, 10, 0, 0) });
            return;
        }
        if (view == null || view.Sheet.Columns.Count == 0)
        {
            if (fingerprint == "empty") return;
            fingerprint = "empty"; headerBadge.Text = ""; ClearContent();
            content.Controls.Add(new Label { Text = "选择表格中的一列，查看或编辑字段配置。", Dock = DockStyle.Top, AutoSize = true, MaximumSize = new Size(240, 0), ForeColor = StudioTheme.Muted, Padding = new Padding(0, 12, 0, 0) });
            return;
        }
        int index = Math.Clamp(view.SelectedColumn, 0, view.Sheet.Columns.Count - 1);
        var column = view.Sheet.Columns[index];
        string state = view.Document.Path + "|" + view.Sheet.Name + "|" + index + "|" + column.IsConfigured + "|" + System.Text.Json.JsonSerializer.Serialize(column) + System.Text.Json.JsonSerializer.Serialize(registry);
        if (!force && (state == fingerprint || ContainsFocus)) return;
        fingerprint = state;
        headerBadge.Text = $"#{index + 1}";
        rendering = true;
        try
        {
            content.SuspendLayout();
            ClearContent();
            var form = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Top, ColumnCount = 1, Padding = Padding.Empty, Margin = Padding.Empty };
            form.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            void Add(Control control, int bottom = 6)
            {
                control.Dock = DockStyle.Top;
                control.Margin = new Padding(0, 0, 0, bottom);
                form.Controls.Add(control);
            }
            void SectionLabel(string text, int top = 12)
            {
                var label = new Label { Text = text, AutoSize = true, ForeColor = StudioTheme.Muted, Font = StudioTheme.UiFont(8.5F), Margin = new Padding(0, top, 0, 5) };
                form.Controls.Add(label);
            }
            void Divider()
            {
                form.Controls.Add(new Panel { Height = 1, Dock = DockStyle.Top, BackColor = StudioTheme.Border, Margin = new Padding(0, 12, 0, 4) });
            }

            string? missing = ColumnConfiguration.Missing(column, registry);
            if (missing != null)
            {
                var banner = new StudioBanner(missing, StudioTheme.WarningBg, StudioTheme.WarningBorder, StudioTheme.WarningText)
                {
                    Dock = DockStyle.Top,
                    Margin = new Padding(0, 4, 0, 8)
                };
                form.Controls.Add(banner);
            }

            SectionLabel("字段名", 4);
            var name = new TextBox { Name = "field_name", Text = column.ColName, BorderStyle = BorderStyle.FixedSingle, Font = StudioTheme.UiFont(9F) };
            Add(name);
            name.Validating += (_, e) =>
            {
                if (rendering || name.Text == column.ColName) return;
                try { view.Document.RenameColumn(view.Sheet, index, name.Text ?? ""); }
                catch (Exception error) { e.Cancel = true; MessageBox.Show(this, error.Message, "无法修改", MessageBoxButtons.OK, MessageBoxIcon.Information); }
            };
            name.ReadOnly = view.Document.ReadOnlyReason != null || view.Sheet.HasTable || view.Sheet.IsProtected;

            SectionLabel("数据类型");
            var type = new ComboBox { Name = "field_type", DropDownStyle = ComboBoxStyle.DropDownList, Font = StudioTheme.UiFont(9F) };
            type.Items.AddRange(UiDefaults.AllowedTypes);
            if (column.ValueType != null && !type.Items.Contains(column.ValueType)) type.Items.Add(column.ValueType);
            type.SelectedItem = column.ValueType;
            Add(type);
            type.SelectedIndexChanged += (_, _) =>
            {
                string? value = type.SelectedItem?.ToString(), before = column.ValueType;
                if (!rendering && (before != value || !column.IsConfigured))
                {
                    bool configured = column.IsConfigured;
                    view.Document.Execute($"修改列 {column.ColName} 类型 → {value}", () => { column.ValueType = value; column.IsConfigured = true; }, () => { column.ValueType = before; column.IsConfigured = configured; });
                }
            };

            if (missing != null)
            {
                var confirm = StudioTheme.Button("确认字段类型", (_, _) =>
                {
                    try { ColumnConfiguration.Confirm(view.Document, column, registry); ShowColumn(view, registry, true); }
                    catch (Exception error) { MessageBox.Show(this, error.Message, "请选择类型", MessageBoxButtons.OK, MessageBoxIcon.Information); }
                }, primary: true, icon: StudioIcon.Check);
                confirm.Name = "confirm_column";
                confirm.AutoSize = false;
                confirm.Height = confirm.GetPreferredSize(Size.Empty).Height;
                confirm.Margin = new Padding(0, 4, 0, 8);
                confirm.Dock = DockStyle.Top;
                form.Controls.Add(confirm);
            }

            SectionLabel("主键序列（-1 为无主键）");
            var primary = new NumericUpDown { Minimum = -1, Maximum = int.MaxValue, Value = Math.Max(-1, column.PkOrder), BorderStyle = BorderStyle.FixedSingle, Font = StudioTheme.UiFont(9F) };
            Add(primary);
            primary.ValueChanged += (_, _) =>
            {
                int before = column.PkOrder, after = (int)primary.Value;
                if (!rendering && before != after) view.Document.Execute($"修改列 {column.ColName} 主键 → {after}", () => column.PkOrder = after, () => column.PkOrder = before);
            };

            Divider();
            SectionLabel("自定义属性（可选）", 8);
            var definitions = registry.Properties.Where(p => p.Scope is "column" or "both").ToList();
            if (definitions.Count == 0) Add(StudioTheme.Label("可在顶部“属性库”中添加自定义项。", true));
            foreach (var definition in definitions)
            {
                var d = definition;
                var include = new CheckBox
                {
                    Text = string.IsNullOrEmpty(d.Title) ? d.Key : $"{d.Title} ({d.Key})",
                    Checked = column.ColProperties?.Any(p => string.Equals(p.PropertyName, d.Key, StringComparison.OrdinalIgnoreCase)) == true,
                    AutoSize = true,
                    ForeColor = StudioTheme.Text,
                    Font = StudioTheme.UiFont(8.75F),
                    Margin = new Padding(0, 8, 0, 4)
                };
                form.Controls.Add(include);

                Control? propertyInput = null;
                string value = column.ColProperties?.FirstOrDefault(p => string.Equals(p.PropertyName, d.Key, StringComparison.OrdinalIgnoreCase))?.PropertyValue ?? d.Default ?? "";
                void Commit(string after)
                {
                    if (rendering || !include.Checked) return;
                    var before = column.ColProperties ?? new List<PropertyKV>();
                    if (before.Any(p => string.Equals(p.PropertyName, d.Key, StringComparison.OrdinalIgnoreCase) && p.PropertyValue == after)) return;
                    var properties = before.Where(p => !string.Equals(p.PropertyName, d.Key, StringComparison.OrdinalIgnoreCase)).Select(p => new PropertyKV(p.PropertyName ?? "", p.PropertyValue ?? "")).ToList();
                    properties.Add(new(d.Key, after));
                    view.Document.Execute($"设置列 {column.ColName} 属性 {d.Key}={after}", () => column.ColProperties = properties, () => column.ColProperties = before);
                }
                switch (d.Kind)
                {
                    case "bool":
                        var check = new CheckBox { Text = "启用 (true)", Checked = string.Equals(value, "true", StringComparison.OrdinalIgnoreCase), AutoSize = true, ForeColor = StudioTheme.Muted, Margin = new Padding(18, 0, 0, 4) };
                        propertyInput = check; check.CheckedChanged += (_, _) => Commit(check.Checked ? "true" : "false"); form.Controls.Add(check); break;
                    case "enum":
                        var select = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Font = StudioTheme.UiFont(9F) };
                        select.Items.AddRange((d.EnumValues ?? new()).ToArray());
                        if (!select.Items.Contains(value)) select.Items.Add(value);
                        propertyInput = select; select.SelectedItem = value; select.SelectedIndexChanged += (_, _) => Commit(select.SelectedItem?.ToString() ?? ""); Add(select); break;
                    case "int":
                        var number = new NumericUpDown { Minimum = int.MinValue, Maximum = int.MaxValue, Value = int.TryParse(value, out int n) ? n : 0, Font = StudioTheme.UiFont(9F) };
                        propertyInput = number; number.ValueChanged += (_, _) => Commit(number.Value.ToString(System.Globalization.CultureInfo.InvariantCulture)); Add(number); break;
                    default:
                        var input = new TextBox { Name = "property:" + d.Key, Text = value, BorderStyle = BorderStyle.FixedSingle, Font = StudioTheme.UiFont(9F) };
                        propertyInput = input; input.Validated += (_, _) => Commit(input.Text); Add(input); break;
                }
                propertyInput!.Enabled = include.Checked;
                include.CheckedChanged += (_, _) =>
                {
                    propertyInput.Enabled = include.Checked;
                    if (rendering) return;
                    if (include.Checked) Commit(propertyInput switch { ComboBox c => c.SelectedItem?.ToString() ?? "", CheckBox c => c.Checked ? "true" : "false", NumericUpDown n => n.Value.ToString(System.Globalization.CultureInfo.InvariantCulture), _ => propertyInput.Text });
                    else
                    {
                        var before = column.ColProperties;
                        var after = (before ?? new()).Where(p => !string.Equals(p.PropertyName, d.Key, StringComparison.OrdinalIgnoreCase)).ToList();
                        view.Document.Execute($"取消列 {column.ColName} 属性 {d.Key}", () => column.ColProperties = after, () => column.ColProperties = before);
                    }
                };
            }

            var extra = (column.ColProperties ?? new()).Where(p => !definitions.Any(d => string.Equals(d.Key, p.PropertyName, StringComparison.OrdinalIgnoreCase))).ToList();
            if (extra.Count > 0)
            {
                Divider();
                SectionLabel("其他已保存属性", 8);
                foreach (var property in extra) Add(StudioTheme.Label(property.PropertyName + " = " + property.PropertyValue, true), 2);
            }

            Divider();
            Add(new Label { Text = "属性独立保存于 .schema.json\nCtrl+S 保存 · Ctrl+Z 撤销", AutoSize = true, Font = StudioTheme.UiFont(8.25F), ForeColor = StudioTheme.SubtleText, Margin = new Padding(0, 6, 0, 0) }, 0);

            content.Controls.Add(form);
        }
        finally { rendering = false; content.ResumeLayout(true); }
    }

    private void ClearContent()
    {
        bool before = rendering; rendering = true;
        try { foreach (Control child in content.Controls.Cast<Control>().ToList()) child.Dispose(); content.Controls.Clear(); }
        finally { rendering = before; }
    }
}

internal sealed class StudioBanner : Control
{
    private readonly Color fill, stroke, fore;
    public StudioBanner(string text, Color bg, Color border, Color fg)
    {
        Text = text; fill = bg; stroke = border; fore = fg;
        Font = StudioTheme.UiFont(8.5F);
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
    }
    protected override void OnSizeChanged(EventArgs e)
    {
        base.OnSizeChanged(e);
        int desired = MeasureHeight(Math.Max(80, Width));
        if (Height != desired) Height = desired;
    }
    public override Size GetPreferredSize(Size proposedSize)
    {
        int w = proposedSize.Width > 40 ? proposedSize.Width : Math.Max(180, Parent?.ClientSize.Width ?? 230);
        return new Size(w, MeasureHeight(w));
    }
    private int MeasureHeight(int width)
    {
        var sz = TextRenderer.MeasureText(Text, Font, new Size(Math.Max(60, width - 20), int.MaxValue), TextFormatFlags.WordBreak | TextFormatFlags.NoPrefix);
        return sz.Height + 16;
    }
    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        var rect = new Rectangle(0, 0, Width - 1, Height - 1);
        if (rect.Width <= 0 || rect.Height <= 0) return;
        using var path = StudioTheme.RoundedRect(rect, 6);
        using var bg = new SolidBrush(fill);
        using var pen = new Pen(stroke, 1f);
        g.FillPath(bg, path);
        g.DrawPath(pen, path);
        TextRenderer.DrawText(g, Text, Font, new Rectangle(10, 8, Math.Max(10, Width - 20), Math.Max(10, Height - 16)), fore,
            TextFormatFlags.WordBreak | TextFormatFlags.Left | TextFormatFlags.Top | TextFormatFlags.NoPrefix);
    }
}
