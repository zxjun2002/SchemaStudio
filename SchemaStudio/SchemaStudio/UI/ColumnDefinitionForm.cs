using SchemaStudio.Models;
using SchemaStudio.Services;

namespace SchemaStudio.UI;

internal sealed class ColumnDefinitionForm : Form
{
    public ColumnSchema? Result { get; private set; }
    public ColumnDefinitionForm(PropertyRegistry registry, IEnumerable<string> existingNames)
    {
        Text = "新增字段";
        ClientSize = new Size(470, Math.Min(720, 340 + ColumnConfiguration.Definitions(registry).Count * 70));
        MinimumSize = new Size(430, 380);
        StartPosition = FormStartPosition.CenterParent;
        MaximizeBox = false; MinimizeBox = false;
        Font = StudioTheme.UiFont(9F);
        AutoScaleMode = AutoScaleMode.Dpi;
        BackColor = StudioTheme.Surface; ForeColor = StudioTheme.Text;
        HandleCreated += (_, _) => StudioTheme.ApplyModernTitleBar(this);

        var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, Padding = new Padding(22, 10, 22, 12), BackColor = StudioTheme.Surface };
        var form = new TableLayoutPanel { Dock = DockStyle.Top, AutoSize = true, ColumnCount = 1 };
        form.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
        void Add(string label, Control input) { form.Controls.Add(StudioTheme.Label(label, true)); input.Dock = DockStyle.Top; input.Margin = new Padding(0, 0, 0, 6); form.Controls.Add(input); }

        var name = new TextBox { Text = "new_field", BorderStyle = BorderStyle.FixedSingle }; Add("字段名", name);
        var type = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList }; type.Items.AddRange(UiDefaults.AllowedTypes); type.SelectedItem = "string"; Add("数据类型", type);
        var primary = new NumericUpDown { Minimum = -1, Maximum = int.MaxValue, Value = -1, BorderStyle = BorderStyle.FixedSingle }; Add("主键序列（-1 为无主键）", primary);

        var inputs = new Dictionary<string, Control>();
        var included = new Dictionary<string, CheckBox>();
        foreach (var definition in ColumnConfiguration.Definitions(registry))
        {
            Control input;
            string value = definition.Default ?? "";
            switch (definition.Kind)
            {
                case "enum":
                    var select = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
                    select.Items.AddRange((definition.EnumValues ?? new()).ToArray());
                    if (value.Length > 0) { if (!select.Items.Contains(value)) select.Items.Add(value); select.SelectedItem = value; }
                    input = select; break;
                case "bool": input = new CheckBox { Text = "启用", Checked = value.Equals("true", StringComparison.OrdinalIgnoreCase), AutoSize = true }; break;
                case "int": input = new NumericUpDown { Minimum = int.MinValue, Maximum = int.MaxValue, Value = int.TryParse(value, out int n) ? n : 0, BorderStyle = BorderStyle.FixedSingle }; break;
                default: input = new TextBox { Text = value, BorderStyle = BorderStyle.FixedSingle }; break;
            }
            inputs[definition.Key] = input;
            var include = new CheckBox { Text = "使用此属性", AutoSize = true, Checked = false, Anchor = AnchorStyles.Left };
            included[definition.Key] = include; input.Enabled = false;
            include.CheckedChanged += (_, _) => input.Enabled = include.Checked;
            var row = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Margin = Padding.Empty };
            row.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            row.Controls.Add(include, 0, 0);
            input.Dock = DockStyle.Fill;
            input.Margin = new Padding(10, 0, 0, 0);
            row.Controls.Add(input, 1, 0);
            Add((string.IsNullOrEmpty(definition.Title) ? definition.Key : definition.Title) + "（可选）", row);
        }
        scroll.Controls.Add(form);

        var actions = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(16, 10, 16, 12), BackColor = StudioTheme.Background };
        actions.Paint += (_, e) => { using var p = new Pen(StudioTheme.Border); e.Graphics.DrawLine(p, 0, 0, actions.Width, 0); };
        var submit = StudioTheme.Button("确认并追加", (_, _) =>
        {
            string field = name.Text.Trim();
            if (field.Length == 0 || existingNames.Any(n => string.Equals(n, field, StringComparison.OrdinalIgnoreCase))) { MessageBox.Show(this, "字段名不能为空或重复。"); return; }
            var values = new List<PropertyKV>();
            foreach (var definition in ColumnConfiguration.Definitions(registry))
            {
                if (!included[definition.Key].Checked) continue;
                var input = inputs[definition.Key];
                string value = input switch { ComboBox c => c.SelectedItem?.ToString() ?? "", CheckBox c => c.Checked ? "true" : "false", NumericUpDown n => n.Value.ToString(System.Globalization.CultureInfo.InvariantCulture), _ => input.Text };
                if (definition.Kind == "enum" && value.Length == 0) { MessageBox.Show(this, "请选择属性：" + definition.Display); input.Focus(); return; }
                values.Add(new(definition.Key, value));
            }
            Result = new ColumnSchema { ColName = field, ValueType = type.SelectedItem!.ToString(), PkOrder = (int)primary.Value, ColProperties = values, IsConfigured = true };
            DialogResult = DialogResult.OK;
        }, primary: true);
        var cancel = StudioTheme.Button("取消", (_, _) => DialogResult = DialogResult.Cancel);
        actions.Controls.Add(submit);
        actions.Controls.Add(cancel);

        Controls.Add(scroll);
        Controls.Add(actions);
        AcceptButton = submit;
        CancelButton = cancel;
    }
}

internal sealed class QuestionButton : Control
{
    public QuestionButton()
    {
        Size = new Size(22, 22); Cursor = Cursors.Hand; TabStop = true;
        AccessibleRole = AccessibleRole.PushButton; AccessibleName = "受保护数据说明";
        SetStyle(ControlStyles.UserPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.SupportsTransparentBackColor, true);
        BackColor = Color.Transparent;
    }
    protected override void OnPaint(PaintEventArgs e)
    {
        e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
        using var pen = new Pen(StudioTheme.Muted, 1.2F);
        e.Graphics.DrawEllipse(pen, 2, 2, 16, 16);
        TextRenderer.DrawText(e.Graphics, "?", StudioTheme.UiFont(8F), new Rectangle(2, 2, 17, 17), StudioTheme.Muted, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
    }
    protected override void OnKeyDown(KeyEventArgs e) { if (e.KeyCode is Keys.Enter or Keys.Space) { OnClick(EventArgs.Empty); e.Handled = true; } base.OnKeyDown(e); }
}
