using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;

namespace SchemaStudio.UI;

internal enum StudioIcon
{
    None,
    FolderOpen,
    FileOpen,
    TableAdd,
    Save,
    Undo,
    Redo,
    Check,
    Sliders,
    Tag,
    More,
    External,
    Refresh,
    RowAdd,
    ColAdd,
    Trash,
    Search,
    PanelRight,
    Table,
    Workbook,
    Csv,
    Folder,
    Lock,
    Warning,
    Info,
    ChevronDown
}

internal static class StudioTheme
{
    // DBX-inspired clean light palette
    public static readonly Color Chrome = Color.White;
    public static readonly Color Background = Color.FromArgb(247, 248, 250);
    public static readonly Color Surface = Color.White;
    public static readonly Color SubtleSurface = Color.FromArgb(243, 245, 248);
    public static readonly Color Hover = Color.FromArgb(238, 242, 247);
    public static readonly Color Pressed = Color.FromArgb(226, 234, 245);
    public static readonly Color Border = Color.FromArgb(229, 233, 240);
    public static readonly Color StrongBorder = Color.FromArgb(207, 214, 226);
    public static readonly Color Text = Color.FromArgb(31, 41, 55);
    public static readonly Color Muted = Color.FromArgb(107, 114, 128);
    public static readonly Color SubtleText = Color.FromArgb(156, 163, 175);
    public static readonly Color Accent = Color.FromArgb(37, 99, 235);
    public static readonly Color AccentHover = Color.FromArgb(29, 78, 216);
    public static readonly Color ActiveTab = Color.FromArgb(219, 234, 250);
    public static readonly Color Selection = Color.FromArgb(235, 243, 254);
    public static readonly Color SheetGreen = Color.FromArgb(22, 163, 106);
    public static readonly Color CsvBlue = Color.FromArgb(37, 99, 235);
    public static readonly Color FolderTint = Color.FromArgb(100, 116, 139);
    public static readonly Color WarningBg = Color.FromArgb(255, 251, 235);
    public static readonly Color WarningBorder = Color.FromArgb(253, 230, 138);
    public static readonly Color WarningText = Color.FromArgb(180, 83, 9);
    public static readonly Color DangerBg = Color.FromArgb(254, 242, 242);
    public static readonly Color DangerText = Color.FromArgb(200, 42, 42);
    public static readonly Color FormulaBg = Color.FromArgb(245, 248, 255);

    public static Font UiFont(float size = 9F, FontStyle style = FontStyle.Regular) => new("Microsoft YaHei UI", size, style);

    public static StudioButton Button(string text, EventHandler click, bool primary = false, StudioIcon icon = StudioIcon.None, bool borderless = false)
    {
        var button = new StudioButton
        {
            Text = text,
            Icon = icon,
            Variant = primary ? StudioButtonVariant.Primary : borderless ? StudioButtonVariant.Ghost : StudioButtonVariant.Secondary,
            Margin = new Padding(0, 0, 6, 0)
        };
        button.Click += click;
        return button;
    }

    public static StudioButton IconButton(StudioIcon icon, string tooltip, EventHandler click, ToolTip? tips = null)
    {
        var button = new StudioButton
        {
            Text = "",
            Icon = icon,
            Variant = StudioButtonVariant.Ghost,
            IconOnly = true,
            AccessibleName = tooltip,
            Margin = new Padding(2, 0, 2, 0)
        };
        button.Click += click;
        tips?.SetToolTip(button, tooltip);
        return button;
    }

    public static int LineHeight(Font font) => TextRenderer.MeasureText("测Ag", font, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix).Height;

    public static void Grid(DataGridView grid)
    {
        var headerFont = UiFont(9F, FontStyle.Regular);
        var cellFont = UiFont(9F);
        int lineH = LineHeight(cellFont);
        grid.BackgroundColor = Surface;
        grid.BorderStyle = BorderStyle.None;
        grid.GridColor = Color.FromArgb(235, 238, 243);
        grid.EnableHeadersVisualStyles = false;
        grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
        grid.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
        grid.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Background,
            ForeColor = Text,
            SelectionBackColor = SubtleSurface,
            SelectionForeColor = Text,
            Font = headerFont,
            Padding = new Padding(10, 0, 6, 0),
            Alignment = DataGridViewContentAlignment.MiddleLeft
        };
        grid.ColumnHeadersHeight = Math.Max(36, lineH + 16);
        grid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
        grid.RowHeadersDefaultCellStyle = new DataGridViewCellStyle
        {
            BackColor = Background,
            ForeColor = SubtleText,
            SelectionBackColor = Selection,
            SelectionForeColor = Accent,
            Font = UiFont(8.5F)
        };
        grid.RowHeadersWidth = Math.Max(52, lineH * 2 + 16);
        grid.DefaultCellStyle = new DataGridViewCellStyle
        {
            ForeColor = Text,
            BackColor = Surface,
            SelectionBackColor = Selection,
            SelectionForeColor = Text,
            Font = cellFont,
            Padding = new Padding(10, 0, 6, 0)
        };
        grid.RowTemplate.Height = Math.Max(30, lineH + 12);
        grid.CellBorderStyle = DataGridViewCellBorderStyle.Single;
        grid.AllowUserToAddRows = false;
        grid.AllowUserToDeleteRows = false;
        grid.AllowUserToOrderColumns = false;
        grid.AutoGenerateColumns = false;
        grid.MultiSelect = true;
        grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
        grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
        EnableDoubleBuffering(grid);
    }

    public static void EnableDoubleBuffering(Control control)
    {
        typeof(Control).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)?.SetValue(control, true);
    }

    public static Label Label(string text, bool muted = false, bool bold = false) => new()
    {
        Text = text,
        AutoSize = true,
        Font = bold ? UiFont(9F, FontStyle.Bold) : UiFont(9F),
        ForeColor = muted ? Muted : Text,
        Margin = new Padding(0, 8, 0, 4)
    };

    public static GraphicsPath RoundedRect(Rectangle rect, int radius)
    {
        var path = new GraphicsPath();
        if (radius <= 0) { path.AddRectangle(rect); return path; }
        int d = Math.Min(radius * 2, Math.Min(rect.Width, rect.Height));
        if (d <= 0) { path.AddRectangle(rect); return path; }
        path.AddArc(rect.X, rect.Y, d, d, 180, 90);
        path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
        path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
        path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }

    public static void ApplyModernTitleBar(Form form)
    {
        if (!OperatingSystem.IsWindowsVersionAtLeast(10, 0, 22000)) return;
        try
        {
            int captionColor = ColorTranslator.ToWin32(Chrome);
            int textColor = ColorTranslator.ToWin32(Text);
            int borderColor = ColorTranslator.ToWin32(Border);
            DwmSetWindowAttribute(form.Handle, 35, ref captionColor, sizeof(int));
            DwmSetWindowAttribute(form.Handle, 36, ref textColor, sizeof(int));
            DwmSetWindowAttribute(form.Handle, 34, ref borderColor, sizeof(int));
        }
        catch { }
    }

    [DllImport("dwmapi.dll")]
    private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

    public static void DrawIcon(Graphics g, StudioIcon icon, Rectangle bounds, Color color)
    {
        if (icon == StudioIcon.None) return;
        var prev = g.SmoothingMode;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        float x = bounds.X, y = bounds.Y, w = bounds.Width, h = bounds.Height;
        float s = Math.Min(w, h) / 16f;
        float ox = x + (w - 16f * s) / 2f;
        float oy = y + (h - 16f * s) / 2f;
        PointF P(float px, float py) => new(ox + px * s, oy + py * s);
        RectangleF R(float rx, float ry, float rw, float rh) => new(ox + rx * s, oy + ry * s, rw * s, rh * s);
        using var pen = new Pen(color, Math.Max(1.25f, 1.35f * s)) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round };
        using var fill = new SolidBrush(color);

        switch (icon)
        {
            case StudioIcon.FolderOpen:
            case StudioIcon.Folder:
                g.DrawLines(pen, new[] { P(1.8f, 3.2f), P(6.2f, 3.2f), P(7.6f, 4.8f), P(14.2f, 4.8f), P(14.2f, 13f), P(1.8f, 13f), P(1.8f, 3.2f) });
                break;
            case StudioIcon.FileOpen:
                g.DrawLines(pen, new[] { P(3f, 2f), P(9.5f, 2f), P(13f, 5.5f), P(13f, 14f), P(3f, 14f), P(3f, 2f) });
                g.DrawLines(pen, new[] { P(9.5f, 2f), P(9.5f, 5.5f), P(13f, 5.5f) });
                g.DrawLine(pen, P(5.5f, 9f), P(10.5f, 9f));
                g.DrawLine(pen, P(5.5f, 11.5f), P(9f, 11.5f));
                break;
            case StudioIcon.TableAdd:
                g.DrawRectangle(pen, R(2f, 2.5f, 11.5f, 10.5f));
                g.DrawLine(pen, P(2f, 6f), P(13.5f, 6f));
                g.DrawLine(pen, P(6f, 6f), P(6f, 13f));
                g.DrawLine(pen, P(10.2f, 8.2f), P(10.2f, 12.2f));
                g.DrawLine(pen, P(8.2f, 10.2f), P(12.2f, 10.2f));
                break;
            case StudioIcon.Save:
                g.DrawLines(pen, new[] { P(2.5f, 2.5f), P(11f, 2.5f), P(13.5f, 5f), P(13.5f, 13.5f), P(2.5f, 13.5f), P(2.5f, 2.5f) });
                g.DrawRectangle(pen, R(5f, 2.5f, 5.2f, 3.5f));
                g.DrawRectangle(pen, R(4.8f, 9f, 6.4f, 4.5f));
                break;
            case StudioIcon.Undo:
                g.DrawArc(pen, R(3.5f, 5.2f, 9f, 7.2f), 210, 175);
                g.DrawLines(pen, new[] { P(5.8f, 3.2f), P(3.2f, 5.6f), P(6.2f, 7.8f) });
                break;
            case StudioIcon.Redo:
                g.DrawArc(pen, R(3.5f, 5.2f, 9f, 7.2f), 155, 175);
                g.DrawLines(pen, new[] { P(10.2f, 3.2f), P(12.8f, 5.6f), P(9.8f, 7.8f) });
                break;
            case StudioIcon.Check:
                g.DrawEllipse(pen, R(2f, 2f, 12f, 12f));
                g.DrawLines(pen, new[] { P(5.2f, 8.2f), P(7.3f, 10.4f), P(11.1f, 5.8f) });
                break;
            case StudioIcon.Sliders:
                g.DrawLine(pen, P(2f, 4.5f), P(14f, 4.5f));
                g.DrawLine(pen, P(2f, 8f), P(14f, 8f));
                g.DrawLine(pen, P(2f, 11.5f), P(14f, 11.5f));
                g.FillEllipse(fill, R(4.2f, 3f, 3f, 3f));
                g.FillEllipse(fill, R(9.2f, 6.5f, 3f, 3f));
                g.FillEllipse(fill, R(5.8f, 10f, 3f, 3f));
                break;
            case StudioIcon.Tag:
                g.DrawLines(pen, new[] { P(2.5f, 3f), P(8f, 3f), P(13.5f, 8.5f), P(8.5f, 13.5f), P(3f, 8f), P(2.5f, 3f) });
                g.FillEllipse(fill, R(4.8f, 5.2f, 2f, 2f));
                break;
            case StudioIcon.More:
                g.FillEllipse(fill, R(2.5f, 6.8f, 2.4f, 2.4f));
                g.FillEllipse(fill, R(6.8f, 6.8f, 2.4f, 2.4f));
                g.FillEllipse(fill, R(11.1f, 6.8f, 2.4f, 2.4f));
                break;
            case StudioIcon.External:
                g.DrawLines(pen, new[] { P(8.5f, 3f), P(13f, 3f), P(13f, 7.5f) });
                g.DrawLine(pen, P(7.5f, 8.5f), P(13f, 3f));
                g.DrawLines(pen, new[] { P(7f, 3.5f), P(3f, 3.5f), P(3f, 13f), P(12.5f, 13f), P(12.5f, 9f) });
                break;
            case StudioIcon.Refresh:
                g.DrawArc(pen, R(2.6f, 2.6f, 10.8f, 10.8f), 35, 290);
                g.DrawLines(pen, new[] { P(11.2f, 2.6f), P(13.4f, 4.6f), P(10.6f, 5.8f) });
                break;
            case StudioIcon.RowAdd:
                g.DrawRectangle(pen, R(2f, 3f, 12f, 4.5f));
                g.DrawLine(pen, P(8f, 9.5f), P(8f, 13.8f));
                g.DrawLine(pen, P(5.8f, 11.6f), P(10.2f, 11.6f));
                break;
            case StudioIcon.ColAdd:
                g.DrawRectangle(pen, R(2.5f, 2.5f, 4.5f, 11f));
                g.DrawLine(pen, P(11.5f, 5.8f), P(11.5f, 10.2f));
                g.DrawLine(pen, P(9.3f, 8f), P(13.7f, 8f));
                break;
            case StudioIcon.Trash:
                g.DrawLine(pen, P(3f, 4.2f), P(13f, 4.2f));
                g.DrawLines(pen, new[] { P(6f, 4.2f), P(6.5f, 2.4f), P(9.5f, 2.4f), P(10f, 4.2f) });
                g.DrawLines(pen, new[] { P(4.3f, 4.2f), P(5f, 13.5f), P(11f, 13.5f), P(11.7f, 4.2f) });
                break;
            case StudioIcon.Search:
                g.DrawEllipse(pen, R(2.5f, 2.5f, 8.5f, 8.5f));
                g.DrawLine(pen, P(9.8f, 9.8f), P(13.6f, 13.6f));
                break;
            case StudioIcon.PanelRight:
                g.DrawRectangle(pen, R(2f, 2.5f, 12f, 11f));
                g.DrawLine(pen, P(9.8f, 2.5f), P(9.8f, 13.5f));
                break;
            case StudioIcon.Table:
                using (var bg = new SolidBrush(Color.FromArgb(32, color)))
                    g.FillRectangle(bg, R(2f, 2.5f, 12f, 3.2f));
                g.DrawRectangle(pen, R(2f, 2.5f, 12f, 11f));
                g.DrawLine(pen, P(2f, 6f), P(14f, 6f));
                g.DrawLine(pen, P(2f, 9.8f), P(14f, 9.8f));
                g.DrawLine(pen, P(6.5f, 6f), P(6.5f, 13.5f));
                break;
            case StudioIcon.Workbook:
                g.DrawRectangle(pen, R(2.2f, 2f, 11.6f, 12f));
                g.DrawLine(pen, P(5.2f, 5.2f), P(10.8f, 10.8f));
                g.DrawLine(pen, P(10.8f, 5.2f), P(5.2f, 10.8f));
                break;
            case StudioIcon.Csv:
                g.DrawLines(pen, new[] { P(3f, 2f), P(9.8f, 2f), P(13f, 5.2f), P(13f, 14f), P(3f, 14f), P(3f, 2f) });
                g.DrawLine(pen, P(5.2f, 8f), P(10.8f, 8f));
                g.DrawLine(pen, P(5.2f, 10.8f), P(10.8f, 10.8f));
                break;
            case StudioIcon.Lock:
                g.DrawRectangle(pen, R(3.5f, 7.2f, 9f, 6.3f));
                g.DrawArc(pen, R(5.2f, 2.8f, 5.6f, 6.2f), 180, 180);
                break;
            case StudioIcon.Warning:
                g.DrawLines(pen, new[] { P(8f, 2.2f), P(14f, 13.2f), P(2f, 13.2f), P(8f, 2.2f) });
                g.DrawLine(pen, P(8f, 6.2f), P(8f, 9.5f));
                g.FillEllipse(fill, R(7.2f, 10.8f, 1.6f, 1.6f));
                break;
            case StudioIcon.Info:
                g.DrawEllipse(pen, R(2f, 2f, 12f, 12f));
                g.FillEllipse(fill, R(7.2f, 4.8f, 1.6f, 1.6f));
                g.DrawLine(pen, P(8f, 7.6f), P(8f, 11.4f));
                break;
            case StudioIcon.ChevronDown:
                g.DrawLines(pen, new[] { P(4.2f, 6.2f), P(8f, 10.2f), P(11.8f, 6.2f) });
                break;
        }
        g.SmoothingMode = prev;
    }
}

internal enum StudioButtonVariant { Ghost, Secondary, Primary }

internal sealed class StudioCard : Panel
{
    public bool BorderTop { get; set; }
    public bool BorderBottom { get; set; }
    public bool BorderAll { get; set; } = true;

    public StudioCard()
    {
        SetStyle(ControlStyles.ResizeRedraw | ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        base.OnPaint(e);
        using var pen = new Pen(StudioTheme.Border, 1f);
        if (BorderAll) e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
        else
        {
            if (BorderTop) e.Graphics.DrawLine(pen, 0, 0, Width, 0);
            if (BorderBottom) e.Graphics.DrawLine(pen, 0, Height - 1, Width, Height - 1);
        }
    }
}

internal sealed class StudioButton : Control, IButtonControl
{
    private bool hovered;
    private bool pressed;
    private StudioIcon icon = StudioIcon.None;
    private StudioButtonVariant variant = StudioButtonVariant.Secondary;
    private bool iconOnly;
    private bool showDropdown;

    public DialogResult DialogResult { get; set; }
    public void NotifyDefault(bool value) { }
    public void PerformClick() { if (CanSelect && Enabled) OnClick(EventArgs.Empty); }

    public StudioIcon Icon { get => icon; set { icon = value; UpdateSize(); Invalidate(); } }
    public StudioButtonVariant Variant { get => variant; set { variant = value; Invalidate(); } }
    public bool IconOnly { get => iconOnly; set { iconOnly = value; UpdateSize(); Invalidate(); } }
    public bool ShowDropdown { get => showDropdown; set { showDropdown = value; UpdateSize(); Invalidate(); } }

    public StudioButton()
    {
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw | ControlStyles.SupportsTransparentBackColor, true);
        Font = StudioTheme.UiFont(9F);
        Cursor = Cursors.Hand;
        AutoSize = true;
        BackColor = Color.Transparent;
        TabStop = true;
    }

    protected override void OnTextChanged(EventArgs e) { base.OnTextChanged(e); UpdateSize(); Invalidate(); }
    protected override void OnFontChanged(EventArgs e) { base.OnFontChanged(e); UpdateSize(); Invalidate(); }
    protected override void OnEnabledChanged(EventArgs e) { base.OnEnabledChanged(e); Cursor = Enabled ? Cursors.Hand : Cursors.Default; Invalidate(); }
    protected override void OnMouseEnter(EventArgs e) { base.OnMouseEnter(e); hovered = true; Invalidate(); }
    protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); hovered = false; pressed = false; Invalidate(); }
    protected override void OnMouseDown(MouseEventArgs e) { base.OnMouseDown(e); if (e.Button == MouseButtons.Left) { pressed = true; Invalidate(); } }
    protected override void OnMouseUp(MouseEventArgs e) { base.OnMouseUp(e); pressed = false; Invalidate(); }
    protected override void OnKeyDown(KeyEventArgs e) { if (e.KeyCode is Keys.Enter or Keys.Space) { OnClick(EventArgs.Empty); e.Handled = true; } base.OnKeyDown(e); }

    private void UpdateSize()
    {
        if (AutoSize && Dock == DockStyle.None) Size = GetPreferredSize(Size.Empty);
    }

    public override Size GetPreferredSize(Size proposedSize)
    {
        var textSize = string.IsNullOrEmpty(Text)
            ? Size.Empty
            : TextRenderer.MeasureText(Text, Font, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
        int lineH = Math.Max(textSize.Height, (int)Math.Ceiling(Font.GetHeight()));
        int h = Math.Max(30, lineH + 14);
        if (iconOnly) return new Size(h, h);
        int padX = variant == StudioButtonVariant.Ghost ? 10 : 14;
        int iconW = icon == StudioIcon.None ? 0 : 16 + (string.IsNullOrEmpty(Text) ? 0 : 6);
        int dropW = showDropdown ? 14 : 0;
        int w = padX * 2 + iconW + textSize.Width + dropW;
        return new Size(Math.Max(variant == StudioButtonVariant.Ghost ? 32 : 72, w), h);
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        var rect = new Rectangle(0, 0, Width - 1, Height - 1);
        if (rect.Width <= 0 || rect.Height <= 0) return;

        Color bg = Color.Transparent;
        Color border = Color.Transparent;
        Color fg = StudioTheme.Text;

        if (!Enabled)
        {
            fg = StudioTheme.SubtleText;
            if (variant == StudioButtonVariant.Primary) { bg = Color.FromArgb(191, 219, 254); fg = Color.White; }
            else if (variant == StudioButtonVariant.Secondary) { bg = StudioTheme.Background; border = StudioTheme.Border; }
        }
        else if (variant == StudioButtonVariant.Primary)
        {
            bg = pressed ? Color.FromArgb(24, 65, 185) : hovered ? StudioTheme.AccentHover : StudioTheme.Accent;
            border = bg;
            fg = Color.White;
        }
        else if (variant == StudioButtonVariant.Secondary)
        {
            bg = pressed ? StudioTheme.Pressed : hovered ? StudioTheme.Background : StudioTheme.Surface;
            border = hovered ? StudioTheme.StrongBorder : StudioTheme.Border;
            fg = StudioTheme.Text;
        }
        else
        {
            bg = pressed ? StudioTheme.Pressed : hovered ? StudioTheme.Hover : Color.Transparent;
            fg = StudioTheme.Text;
        }

        using var path = StudioTheme.RoundedRect(rect, 6);
        if (bg != Color.Transparent) { using var brush = new SolidBrush(bg); g.FillPath(brush, path); }
        if (border != Color.Transparent) { using var pen = new Pen(border, 1f); g.DrawPath(pen, path); }

        int iconSize = 15;
        if (iconOnly)
        {
            var iconRect = new Rectangle((Width - iconSize) / 2, (Height - iconSize) / 2, iconSize, iconSize);
            StudioTheme.DrawIcon(g, icon, iconRect, fg);
            return;
        }

        var textSize = string.IsNullOrEmpty(Text)
            ? Size.Empty
            : TextRenderer.MeasureText(Text, Font, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
        int contentW = (icon == StudioIcon.None ? 0 : iconSize + (string.IsNullOrEmpty(Text) ? 0 : 6)) + textSize.Width + (showDropdown ? 14 : 0);
        int x = Math.Max(8, (Width - contentW) / 2);

        if (icon != StudioIcon.None)
        {
            StudioTheme.DrawIcon(g, icon, new Rectangle(x, (Height - iconSize) / 2, iconSize, iconSize), variant == StudioButtonVariant.Primary ? Color.White : Enabled ? StudioTheme.Muted : StudioTheme.SubtleText);
            x += iconSize + 6;
        }
        if (!string.IsNullOrEmpty(Text))
        {
            TextRenderer.DrawText(g, Text, Font, new Rectangle(x, 0, Math.Max(1, Width - x - (showDropdown ? 16 : 4)), Height), fg,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
            x += textSize.Width;
        }
        if (showDropdown)
        {
            StudioTheme.DrawIcon(g, StudioIcon.ChevronDown, new Rectangle(x + 2, (Height - 12) / 2, 12, 12), fg);
        }
    }
}

internal sealed class StudioSearchBox : Panel
{
    public TextBox Input { get; }
    private readonly Label clear;
    private bool focused;

    public string Query { get => Input.Text; set => Input.Text = value; }
    public event EventHandler? QueryChanged;

    public StudioSearchBox(string placeholder = "搜索配置表...")
    {
        Height = 32;
        Padding = new Padding(28, 5, 24, 5);
        BackColor = StudioTheme.Surface;
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);

        Input = new TextBox
        {
            BorderStyle = BorderStyle.None,
            Dock = DockStyle.Fill,
            BackColor = StudioTheme.Surface,
            ForeColor = StudioTheme.Text,
            Font = StudioTheme.UiFont(9F),
            PlaceholderText = placeholder
        };
        Input.GotFocus += (_, _) => { focused = true; Invalidate(); };
        Input.LostFocus += (_, _) => { focused = false; Invalidate(); };

        clear = new Label
        {
            Text = "×",
            Visible = false,
            Cursor = Cursors.Hand,
            ForeColor = StudioTheme.Muted,
            BackColor = Color.Transparent,
            TextAlign = ContentAlignment.MiddleCenter,
            Size = new Size(20, 20),
            Anchor = AnchorStyles.Top | AnchorStyles.Right
        };
        Input.TextChanged += (_, _) => { clear.Visible = Input.Text.Length > 0; QueryChanged?.Invoke(this, EventArgs.Empty); };
        clear.Click += (_, _) => { Input.Text = ""; Input.Focus(); };
        clear.MouseEnter += (_, _) => clear.ForeColor = StudioTheme.Text;
        clear.MouseLeave += (_, _) => clear.ForeColor = StudioTheme.Muted;

        Controls.Add(Input);
        Controls.Add(clear);
        Resize += (_, _) => PositionInner();
        PositionInner();
    }

    private void PositionInner()
    {
        clear.Location = new Point(Width - 23, (Height - clear.Height) / 2);
        int top = Math.Max(4, (Height - Input.PreferredHeight) / 2);
        Padding = new Padding(28, top, 24, Math.Max(2, Height - top - Input.PreferredHeight));
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        var rect = new Rectangle(0, 0, Width - 1, Height - 1);
        using var path = StudioTheme.RoundedRect(rect, 6);
        using var bg = new SolidBrush(StudioTheme.Surface);
        g.FillPath(bg, path);
        using var pen = new Pen(focused ? StudioTheme.Accent : StudioTheme.Border, 1f);
        g.DrawPath(pen, path);
        StudioTheme.DrawIcon(g, StudioIcon.Search, new Rectangle(8, (Height - 14) / 2, 14, 14), focused ? StudioTheme.Accent : StudioTheme.SubtleText);
    }
}

internal sealed class StudioTabItem
{
    public string Title { get; set; } = "";
    public bool IsDirty { get; set; }
    public StudioIcon Icon { get; set; } = StudioIcon.Table;
    public Color IconColor { get; set; } = StudioTheme.SheetGreen;
    public Control Content { get; }
    public object? Tag { get; set; }

    public StudioTabItem(Control content, string title = "")
    {
        Content = content;
        Title = title;
        Tag = content;
    }
}

internal sealed class StudioTabStrip : Control
{
    private readonly List<StudioTabItem> items = new();
    private readonly List<Rectangle> tabRects = new();
    private readonly List<Rectangle> closeRects = new();
    private int selectedIndex = -1;
    private int hoverTab = -1;
    private int hoverClose = -1;

    public IReadOnlyList<StudioTabItem> Tabs => items;
    public int Count => items.Count;
    public int SelectedIndex
    {
        get => selectedIndex;
        set
        {
            if (value == selectedIndex || value < -1 || value >= items.Count) return;
            if (Selecting != null)
            {
                var args = new System.ComponentModel.CancelEventArgs();
                Selecting(this, args);
                if (args.Cancel) return;
            }
            selectedIndex = value;
            LayoutTabs();
            Invalidate();
            SelectedIndexChanged?.Invoke(this, EventArgs.Empty);
        }
    }
    public StudioTabItem? SelectedTab => selectedIndex >= 0 && selectedIndex < items.Count ? items[selectedIndex] : null;

    public event EventHandler<System.ComponentModel.CancelEventArgs>? Selecting;
    public event EventHandler? SelectedIndexChanged;
    public event Action<StudioTabItem>? CloseRequested;

    public StudioTabStrip()
    {
        Height = 38;
        Font = StudioTheme.UiFont(9F);
        BackColor = StudioTheme.Background;
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
    }

    public void AddTab(StudioTabItem item)
    {
        items.Add(item);
        LayoutTabs();
        SelectedIndex = items.Count - 1;
        Invalidate();
    }

    public void RemoveTab(StudioTabItem item)
    {
        int idx = items.IndexOf(item);
        if (idx < 0) return;
        items.RemoveAt(idx);
        if (items.Count == 0) selectedIndex = -1;
        else if (selectedIndex >= items.Count) selectedIndex = items.Count - 1;
        else if (selectedIndex > idx) selectedIndex--;
        LayoutTabs();
        Invalidate();
        SelectedIndexChanged?.Invoke(this, EventArgs.Empty);
    }

    public void SelectTab(StudioTabItem item)
    {
        int idx = items.IndexOf(item);
        if (idx >= 0) SelectedIndex = idx;
    }

    public void RefreshTabs()
    {
        LayoutTabs();
        Invalidate();
    }

    protected override void OnResize(EventArgs e) { base.OnResize(e); LayoutTabs(); }

    private void LayoutTabs()
    {
        tabRects.Clear();
        closeRects.Clear();
        if (items.Count == 0) return;

        int x = 6;
        int available = Math.Max(120, Width - 16);
        int maxPerTab = Math.Max(110, available / Math.Max(1, items.Count));
        foreach (var item in items)
        {
            string display = item.Title + (item.IsDirty ? " ●" : "");
            int textW = TextRenderer.MeasureText(display, Font, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix).Width;
            int desired = Math.Clamp(textW + 56, 118, 250);
            int width = items.Count * 118 <= available ? Math.Min(desired, Math.Max(120, maxPerTab)) : Math.Max(90, available / items.Count);
            var r = new Rectangle(x, 4, width, Height - 5);
            tabRects.Add(r);
            closeRects.Add(new Rectangle(r.Right - 24, r.Top + (r.Height - 18) / 2, 18, 18));
            x += width + 2;
        }
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        int newTab = -1, newClose = -1;
        for (int i = 0; i < tabRects.Count; i++)
        {
            if (tabRects[i].Contains(e.Location))
            {
                newTab = i;
                if (closeRects[i].Contains(e.Location)) newClose = i;
                break;
            }
        }
        if (newTab != hoverTab || newClose != hoverClose)
        {
            hoverTab = newTab; hoverClose = newClose;
            Cursor = newTab >= 0 ? Cursors.Hand : Cursors.Default;
            Invalidate();
        }
    }

    protected override void OnMouseLeave(EventArgs e)
    {
        base.OnMouseLeave(e);
        hoverTab = hoverClose = -1;
        Cursor = Cursors.Default;
        Invalidate();
    }

    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        for (int i = 0; i < tabRects.Count; i++)
        {
            if (!tabRects[i].Contains(e.Location)) continue;
            if (e.Button == MouseButtons.Middle || (e.Button == MouseButtons.Left && closeRects[i].Contains(e.Location)))
            {
                CloseRequested?.Invoke(items[i]);
                return;
            }
            if (e.Button == MouseButtons.Left)
            {
                SelectedIndex = i;
                return;
            }
        }
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var stripBg = new SolidBrush(StudioTheme.Background);
        g.FillRectangle(stripBg, ClientRectangle);
        using var borderPen = new Pen(StudioTheme.Border, 1f);
        g.DrawLine(borderPen, 0, Height - 1, Width, Height - 1);

        for (int i = 0; i < items.Count && i < tabRects.Count; i++)
        {
            var item = items[i];
            var rect = tabRects[i];
            var close = closeRects[i];
            bool active = i == selectedIndex;
            bool hovered = i == hoverTab;

            using var path = StudioTheme.RoundedRect(new Rectangle(rect.X, rect.Y, rect.Width, rect.Height + 4), 6);
            Color fill = active ? StudioTheme.ActiveTab : hovered ? StudioTheme.Hover : Color.Transparent;
            if (fill != Color.Transparent)
            {
                using var b = new SolidBrush(fill);
                g.FillPath(b, path);
            }
            if (active)
            {
                using var accentPen = new Pen(StudioTheme.Accent, 2f);
                g.DrawLine(accentPen, rect.Left + 1, Height - 2, rect.Right - 1, Height - 2);
            }
            else if (i < items.Count - 1 && i + 1 != selectedIndex)
            {
                g.DrawLine(borderPen, rect.Right + 1, rect.Top + 7, rect.Right + 1, rect.Bottom - 7);
            }

            StudioTheme.DrawIcon(g, item.Icon, new Rectangle(rect.Left + 10, rect.Top + (rect.Height - 14) / 2, 14, 14), item.IconColor);

            int textRight = close.Left - 2;
            if (item.IsDirty)
            {
                using var dot = new SolidBrush(StudioTheme.Accent);
                g.FillEllipse(dot, textRight - 8, rect.Top + (rect.Height - 6) / 2, 6, 6);
                textRight -= 12;
            }
            var textRect = new Rectangle(rect.Left + 29, rect.Top, Math.Max(10, textRight - rect.Left - 29), rect.Height);
            TextRenderer.DrawText(g, item.Title, Font, textRect, active ? StudioTheme.Text : StudioTheme.Muted,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);

            if (i == hoverClose)
            {
                using var closeBg = new SolidBrush(Color.FromArgb(30, 0, 0, 0));
                g.FillEllipse(closeBg, close);
            }
            TextRenderer.DrawText(g, "×", Font, close, i == hoverClose ? StudioTheme.Text : StudioTheme.Muted,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
        }
    }
}

internal sealed class StudioSegmentedControl : Control
{
    private readonly string[] labels;
    private int selectedIndex;
    private int hoverIndex = -1;
    private Rectangle[] itemRects = Array.Empty<Rectangle>();

    public int SelectedIndex
    {
        get => selectedIndex;
        set
        {
            if (value == selectedIndex || value < 0 || value >= labels.Length) return;
            selectedIndex = value;
            Invalidate();
            SelectedIndexChanged?.Invoke(this, EventArgs.Empty);
        }
    }

    public event EventHandler? SelectedIndexChanged;

    public StudioSegmentedControl(params string[] items)
    {
        labels = items;
        Font = StudioTheme.UiFont(8.75F);
        Cursor = Cursors.Hand;
        SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.ResizeRedraw, true);
        Size = GetPreferredSize(Size.Empty);
    }

    public override Size GetPreferredSize(Size proposedSize)
    {
        int total = 6;
        int lineH = (int)Math.Ceiling(Font.GetHeight());
        foreach (var label in labels)
        {
            int w = TextRenderer.MeasureText(label, Font, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix).Width;
            total += w + 22;
        }
        return new Size(total, Math.Max(28, lineH + 12));
    }

    private void ComputeRects()
    {
        itemRects = new Rectangle[labels.Length];
        int x = 3;
        for (int i = 0; i < labels.Length; i++)
        {
            int w = TextRenderer.MeasureText(labels[i], Font, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix).Width + 22;
            itemRects[i] = new Rectangle(x, 3, w, Height - 6);
            x += w;
        }
    }

    protected override void OnMouseMove(MouseEventArgs e)
    {
        base.OnMouseMove(e);
        ComputeRects();
        int idx = -1;
        for (int i = 0; i < itemRects.Length; i++) if (itemRects[i].Contains(e.Location)) { idx = i; break; }
        if (idx != hoverIndex) { hoverIndex = idx; Invalidate(); }
    }

    protected override void OnMouseLeave(EventArgs e) { base.OnMouseLeave(e); hoverIndex = -1; Invalidate(); }
    protected override void OnMouseDown(MouseEventArgs e)
    {
        base.OnMouseDown(e);
        if (e.Button != MouseButtons.Left) return;
        ComputeRects();
        for (int i = 0; i < itemRects.Length; i++) if (itemRects[i].Contains(e.Location)) { SelectedIndex = i; return; }
    }

    protected override void OnPaint(PaintEventArgs e)
    {
        ComputeRects();
        var g = e.Graphics;
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var outer = StudioTheme.RoundedRect(new Rectangle(0, 0, Width - 1, Height - 1), 6);
        using var bg = new SolidBrush(StudioTheme.SubtleSurface);
        g.FillPath(bg, outer);

        for (int i = 0; i < labels.Length; i++)
        {
            var r = itemRects[i];
            bool active = i == selectedIndex;
            if (active || i == hoverIndex)
            {
                using var p = StudioTheme.RoundedRect(r, 5);
                using var b = new SolidBrush(active ? StudioTheme.Surface : StudioTheme.Hover);
                g.FillPath(b, p);
                if (active)
                {
                    using var border = new Pen(StudioTheme.Border, 1f);
                    g.DrawPath(border, p);
                }
            }
            TextRenderer.DrawText(g, labels[i], Font, r, active ? StudioTheme.Text : StudioTheme.Muted,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix);
        }
    }
}
