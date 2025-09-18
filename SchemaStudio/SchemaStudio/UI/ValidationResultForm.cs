using SchemaStudio.Models;
using System.Text;

namespace SchemaStudio.UI
{
    public class ValidationResultForm : Form
    {
        private readonly DataGridView grid;
        private readonly Button btnExport, btnClose;

        public ValidationResultForm(string tableName, List<ValidationError> errors)
        {
            Text = errors.Count == 0 ? $"✅ {tableName}：未发现类型问题" : $"❌ {tableName}：共 {errors.Count} 处类型不匹配";
            Width = 900; Height = 520;

            grid = new DataGridView { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false, AutoGenerateColumns = false };
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "行号", DataPropertyName = nameof(ValidationError.Row), Width = 70 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "列名", DataPropertyName = nameof(ValidationError.Column), Width = 160 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "期望类型", DataPropertyName = nameof(ValidationError.Expected), Width = 100 });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "实际值", DataPropertyName = nameof(ValidationError.Value), AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill });
            grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "原因", DataPropertyName = nameof(ValidationError.Message), Width = 220 });
            grid.DataSource = errors;

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 44 };
            btnExport = new Button { Text = "导出CSV", Left = 10, Top = 8, Width = 90, Enabled = errors.Count > 0 };
            btnClose  = new Button { Text = "关闭", Left = 110, Top = 8, Width = 80 };

            btnExport.Click += (_, __) =>
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
            };
            btnClose.Click += (_, __) => Close();

            bottom.Controls.Add(btnExport);
            bottom.Controls.Add(btnClose);
            Controls.Add(grid);
            Controls.Add(bottom);
            if (errors.Count == 0) btnExport.Enabled = false;
        }
    }
}
