using SchemaStudio.Models;
using System.Text;

namespace SchemaStudio.UI
{
    public class ValidationResultForm : Form
    {
        private readonly DataGridView grid;

        public ValidationResultForm(string tableName, List<ValidationError> errors)
        {
            Text = errors.Count == 0 ? $"{tableName}：未发现类型问题" : $"{tableName}：共 {errors.Count} 处类型不匹配";
            ClientSize = new Size(900, 520);
            MinimumSize = new Size(640, 360);
            StartPosition = FormStartPosition.CenterParent;
            Font = StudioTheme.UiFont(9F);
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = StudioTheme.Surface;
            ForeColor = StudioTheme.Text;
            HandleCreated += (_, _) => StudioTheme.ApplyModernTitleBar(this);

            grid = new DataGridView { Dock = DockStyle.Fill };
            StudioTheme.Grid(grid);
            grid.ReadOnly = true;
            grid.RowHeadersVisible = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "行号", DataPropertyName = nameof(ValidationError.Row), Width = 70 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "列名", DataPropertyName = nameof(ValidationError.Column), Width = 160 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "期望类型", DataPropertyName = nameof(ValidationError.Expected), Width = 110 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "实际值", DataPropertyName = nameof(ValidationError.Value), AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "原因", DataPropertyName = nameof(ValidationError.Message), Width = 220 });
            grid.DataSource = errors;

            var bottom = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(14, 10, 14, 12), BackColor = StudioTheme.Background };
            bottom.Paint += (_, e) => { using var p = new Pen(StudioTheme.Border); e.Graphics.DrawLine(p, 0, 0, bottom.Width, 0); };

            var btnClose = StudioTheme.Button("关闭", (_, _) => Close());
            var btnExport = StudioTheme.Button("导出 CSV", (_, _) =>
            {
                using var sfd = new SaveFileDialog { Filter = "CSV 文件|*.csv", FileName = $"{tableName}.type_errors.csv" };
                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    using var sw = new StreamWriter(sfd.FileName, false, new UTF8Encoding(false));
                    sw.WriteLine("Row,Column,Expected,Value,Message");
                    foreach (ValidationError e in (List<ValidationError>)grid.DataSource!)
                    {
                        string esc(string s) => "\"" + s.Replace("\"", "\"\"") + "\"";
                        sw.WriteLine($"{e.Row},{esc(e.Column)},{esc(e.Expected)},{esc(e.Value)},{esc(e.Message)}");
                    }
                    MessageBox.Show("已导出。");
                }
            }, primary: true);
            btnExport.Enabled = errors.Count > 0;

            bottom.Controls.Add(btnClose);
            bottom.Controls.Add(btnExport);
            Controls.Add(grid);
            Controls.Add(bottom);
            CancelButton = btnClose;
        }
    }
}
