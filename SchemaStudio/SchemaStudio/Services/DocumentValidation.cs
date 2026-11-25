using SchemaStudio.Models;

namespace SchemaStudio.Services;

public static class DocumentValidation
{
    public static List<ValidationError> Validate(EditorSheet sheet, PropertyRegistry? registry = null)
    {
        var errors = new List<ValidationError>();
        if (registry != null)
            foreach (var column in sheet.Columns)
                if (ColumnConfiguration.Missing(column, registry) is { } missing)
                    errors.Add(new() { Row = 1, Column = column.ColName ?? "", Message = missing, Expected = "已确认的数据类型" });
        var primary = sheet.Columns.Select((c, i) => (Column: c, Index: i)).Where(c => c.Column.PkOrder >= 0).OrderBy(c => c.Column.PkOrder).ToList();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        if (primary.GroupBy(p => p.Column.PkOrder).Any(g => g.Count() > 1))
            errors.Add(new() { Row = 1, Column = "主键", Message = "主键序列重复，请为联合主键设置不同的序号。" });
        for (int row = 1; row <= sheet.RowCount; row++)
        {
            bool hasData = Enumerable.Range(0, sheet.Columns.Count).Any(col => sheet.Cell(row, col).Text.Length > 0);
            if (!hasData) continue;
            for (int col = 0; col < sheet.Columns.Count; col++)
            {
                var column = sheet.Columns[col];
                var cell = sheet.Cell(row, col);
                if (cell.IsFormula)
                {
                    if (!cell.HasCachedValue || new[] { "#REF!", "#VALUE!", "#DIV/0!", "#N/A", "#NAME?", "#NUM!", "#NULL!" }.Contains(cell.Cached))
                        errors.Add(new() { Row = row + 1, Column = column.ColName ?? "", Value = cell.Cached, Message = !cell.HasCachedValue ? "公式没有缓存结果，请在 Excel / WPS 中重算并保存。" : "公式缓存是错误值，请在 Excel / WPS 中检查。" });
                    continue; // Never treat a saved cache as a fresh calculation.
                }
                if (!TypeInference.IsValid(cell.Text, column.ValueType ?? "string", out var reason))
                    errors.Add(new() { Row = row + 1, Column = column.ColName ?? "", Expected = column.ValueType ?? "string", Value = cell.Text, Message = reason });
            }
            if (primary.Count == 0 || primary.Any(p => sheet.Cell(row, p.Index).IsFormula)) continue;
            var parts = primary.Select(p => sheet.Cell(row, p.Index).Text).ToList();
            string key = System.Text.Json.JsonSerializer.Serialize(parts);
            if (parts.Any(string.IsNullOrWhiteSpace) || !seen.Add(key))
                errors.Add(new() { Row = row + 1, Column = string.Join(" + ", primary.Select(p => p.Column.ColName)), Expected = "唯一且非空", Value = string.Join(" / ", parts), Message = "主键为空或重复。" });
        }
        return errors;
    }
}
