using System.Text;
using SchemaStudio.Services;

namespace SchemaStudio.Models;

public sealed record EditorCell(string Text, string Cached = "", string Kind = "string", string? Protection = null, bool HasCachedValue = true)
{
    public bool IsFormula => Kind == "formula";
    public string Display => IsFormula ? Cached : Text;
}

public sealed record HistoryEntry(int TargetCursor, string Label, DateTime? Timestamp, bool IsCurrent, bool IsUndone);

public sealed class EditorSheet
{
    public string Name { get; init; } = "";
    public string? PartPath { get; init; }
    public List<ColumnSchema> Columns { get; set; } = new();
    public List<PropertyKV> Properties { get; set; } = new();
    // Keys are Excel cell addresses, including the field-name row.
    public Dictionary<string, EditorCell> Cells { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, EditorCell> OriginalCells { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public List<string> ProtectedRanges { get; } = new();
    public bool HasTable { get; set; }
    public bool IsProtected { get; set; }
    public string? SchemaIssue { get; set; }
    public int RowCount { get; set; }
    public int OriginalRowCount { get; set; }
    public int OriginalColumnCount { get; set; }

    public EditorCell Cell(int row, int column) => Cells.GetValueOrDefault(Address(row, column)) ?? new("");
    public string? ProtectionAt(int row, int column)
    {
        if (IsProtected) return "此工作表受保护，请在 Excel 中解除保护后再编辑。";
        var protection = Cell(row, column).Protection;
        if (protection != null) return protection;
        if (Cell(row, column).IsFormula) return "这是公式单元格，显示文件中上次保存的计算结果。公式受保护，不能在此编辑。";
        return ProtectedRanges.Any(range => Contains(range, row, column))
            ? "数组、共享公式或合并区域受保护，请在 Excel 中编辑。" : null;
    }

    public static string Address(int row, int column)
    {
        string letters = "";
        for (int n = column + 1; n > 0; n = (n - 1) / 26) letters = (char)('A' + (n - 1) % 26) + letters;
        return letters + (row + 1);
    }

    public static (int Row, int Column) Position(string address)
    {
        int col = 0, i = 0;
        address = address.Replace("$", "");
        while (i < address.Length && char.IsLetter(address[i])) col = col * 26 + char.ToUpperInvariant(address[i++]) - 'A' + 1;
        return (int.Parse(address[i..]) - 1, col - 1);
    }

    private static bool Contains(string range, int row, int col)
    {
        var ends = range.Split(':');
        var first = Position(ends[0]);
        var last = Position(ends[^1]);
        return row >= first.Row && row <= last.Row && col >= first.Column && col <= last.Column;
    }
}

public sealed class EditorDocument
{
    public const int MaxHistory = 100;
    private sealed record Change(string Label, DateTime Timestamp, Action Redo, Action Undo, int Before, int After);
    private readonly List<Change> history = new();
    private int cursor, trimmedCount, sequence, revision, savedRevision;
    public string Path { get; init; } = "";
    public bool IsCsv => System.IO.Path.GetExtension(Path).Equals(".csv", StringComparison.OrdinalIgnoreCase);
    public bool IsLegacy => System.IO.Path.GetExtension(Path).Equals(".xls", StringComparison.OrdinalIgnoreCase);
    public string? ReadOnlyReason { get; set; }
    public bool HasExternalLinks { get; set; }
    public List<EditorSheet> Sheets { get; } = new();
    public List<string> Notices { get; } = new();
    public Encoding Encoding { get; set; } = new UTF8Encoding(false);
    public string Delimiter { get; set; } = ",";
    public string NewLine { get; set; } = Environment.NewLine;
    public byte[] SourceBytes { get; set; } = Array.Empty<byte>();
    public byte[]? SchemaBytes { get; set; }
    public WorkbookSchema? SchemaEnvelope { get; set; }
    public TableSchema? TableEnvelope { get; set; }
    public bool IsDirty => revision != savedRevision;
    public bool CanUndo => cursor > 0;
    public bool CanRedo => cursor < history.Count;
    public int HistoryCount => history.Count;
    public int HistoryCursor => cursor;
    public bool HistoryTrimmed => trimmedCount > 0;
    public event EventHandler? Changed;

    public IReadOnlyList<HistoryEntry> GetHistory()
    {
        var items = new List<HistoryEntry>(history.Count + 1);
        for (int i = history.Count; i >= 1; i--)
        {
            var entry = history[i - 1];
            items.Add(new HistoryEntry(i, entry.Label, entry.Timestamp, i == cursor, i > cursor));
        }
        string baseline = trimmedCount > 0 ? "最早可回退状态（已超过 100 条上限）" : "初始状态（打开时）";
        items.Add(new HistoryEntry(0, baseline, null, cursor == 0, false));
        return items;
    }

    public void Execute(string label, Action apply, Action revert)
    {
        apply();
        if (cursor < history.Count)
            history.RemoveRange(cursor, history.Count - cursor);
        var command = new Change(label, DateTime.Now, apply, revert, revision, ++sequence);
        history.Add(command);
        cursor = history.Count;
        revision = command.After;
        if (history.Count > MaxHistory)
        {
            history.RemoveAt(0);
            cursor--;
            trimmedCount++;
        }
        Changed?.Invoke(this, EventArgs.Empty);
    }

    public void Undo()
    {
        if (cursor > 0) JumpTo(cursor - 1);
    }

    public void Redo()
    {
        if (cursor < history.Count) JumpTo(cursor + 1);
    }

    public void JumpTo(int targetCursor)
    {
        targetCursor = Math.Clamp(targetCursor, 0, history.Count);
        if (targetCursor == cursor) return;
        while (cursor > targetCursor)
        {
            cursor--;
            var cmd = history[cursor];
            cmd.Undo();
            revision = cmd.Before;
        }
        while (cursor < targetCursor)
        {
            var cmd = history[cursor];
            cmd.Redo();
            revision = cmd.After;
            cursor++;
        }
        Changed?.Invoke(this, EventArgs.Empty);
    }

    public void MarkSaved()
    {
        savedRevision = revision;
        Changed?.Invoke(this, EventArgs.Empty);
    }

    public void SetCells(EditorSheet sheet, IReadOnlyList<(int Row, int Column, string Text)> edits)
    {
        if (ReadOnlyReason != null) throw new InvalidOperationException(ReadOnlyReason);
        var changes = new List<(string Address, EditorCell? Before, EditorCell After)>();
        foreach (var (row, col, text) in edits)
        {
            if (row < 1 || col < 0 || col >= sheet.Columns.Count) throw new InvalidOperationException("请选择有效的数据单元格。");
            if (row >= 1_048_576 || col >= 16_384) throw new InvalidOperationException("超出表格的行列范围。");
            string address = EditorSheet.Address(row, col);
            var before = sheet.Cells.GetValueOrDefault(address);
            if ((before?.Text ?? "") == text) continue;
            if (sheet.ProtectionAt(row, col) is { } protection) throw new InvalidOperationException(protection);
            // Text beginning with '=' remains text. This editor never creates or rewrites formulas.
            changes.Add((address, before, new(text, Kind: "string")));
        }
        if (changes.Count == 0) return;
        int beforeRows = sheet.RowCount;
        string label = changes.Count == 1 ? $"编辑单元格 {changes[0].Address}" : $"批量编辑 {changes.Count} 个单元格";
        Execute(label, () =>
        {
            foreach (var change in changes) sheet.Cells[change.Address] = change.After;
            sheet.RowCount = Math.Max(beforeRows, edits.Max(e => e.Row));
        }, () =>
        {
            foreach (var change in changes)
                if (change.Before == null) sheet.Cells.Remove(change.Address); else sheet.Cells[change.Address] = change.Before;
            sheet.RowCount = beforeRows;
        });
    }

    public void RenameColumn(EditorSheet sheet, int index, string name)
    {
        name = name.Trim();
        if (name.Length == 0 || sheet.Columns.Where((_, i) => i != index).Any(c => string.Equals(c.ColName, name, StringComparison.OrdinalIgnoreCase)))
            throw new InvalidOperationException("字段名不能为空或重复。");
        if (ReadOnlyReason != null || sheet.IsProtected || sheet.HasTable)
            throw new InvalidOperationException(ReadOnlyReason ?? "受保护工作表或 Excel 结构化表的字段名请在 Excel 中修改。");
        string? before = sheet.Columns[index].ColName;
        var oldCell = sheet.Cell(0, index);
        if (oldCell.IsFormula || oldCell.Protection != null) throw new InvalidOperationException("此表头受公式保护。");
        if (before == name) return;
        Execute($"重命名列 {before} → {name}", () => { sheet.Columns[index].ColName = name; sheet.Cells[EditorSheet.Address(0, index)] = new(name); },
            () => { sheet.Columns[index].ColName = before; sheet.Cells[EditorSheet.Address(0, index)] = oldCell; });
    }

    public void AddColumn(EditorSheet sheet, ColumnSchema column)
    {
        string name = (column.ColName ?? "").Trim();
        if (ReadOnlyReason != null || sheet.IsProtected || sheet.HasTable) throw new InvalidOperationException(ReadOnlyReason ?? "此工作表的结构请在 Excel 中修改。");
        if (!column.IsConfigured || string.IsNullOrWhiteSpace(column.ValueType)) throw new InvalidOperationException("新增列前请选择数据类型。");
        if (name.Length == 0 || sheet.Columns.Any(c => string.Equals(c.ColName, name, StringComparison.OrdinalIgnoreCase))) throw new InvalidOperationException("字段名不能为空或重复。");
        if (sheet.Columns.Count >= 16_384) throw new InvalidOperationException("已达到最大列数。");
        int index = sheet.Columns.Count;
        column.ColName = name; column.ColSort = index;
        Execute($"追加列 {name} ({column.ValueType})", () => { sheet.Columns.Add(column); sheet.Cells[EditorSheet.Address(0, index)] = new(name); },
            () => { sheet.Columns.RemoveAt(index); sheet.Cells.Remove(EditorSheet.Address(0, index)); });
    }

    public void AddRow(EditorSheet sheet)
    {
        if (ReadOnlyReason != null || sheet.IsProtected) throw new InvalidOperationException(ReadOnlyReason ?? "工作表受保护。");
        if (sheet.Columns.Count == 0) throw new InvalidOperationException("请先追加一个字段。");
        if (sheet.RowCount >= 1_048_575) throw new InvalidOperationException("已达到最大行数。");
        int before = sheet.RowCount;
        Execute($"追加第 {before + 2} 行", () => sheet.RowCount = before + 1, () => sheet.RowCount = before);
    }

    public void DeleteRow(EditorSheet sheet, int row)
    {
        if (!IsCsv) throw new InvalidOperationException("Excel 数据为只读，请在 Excel 中编辑。");
        if (row < 1 || row > sheet.RowCount) return;
        var before = sheet.Cells;
        var after = new Dictionary<string, EditorCell>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, cell) in before)
        {
            var p = EditorSheet.Position(key);
            if (p.Row != row) after[EditorSheet.Address(p.Row > row ? p.Row - 1 : p.Row, p.Column)] = cell;
        }
        int count = sheet.RowCount;
        Execute($"删除第 {row + 1} 行", () => { sheet.Cells = after; sheet.RowCount = count - 1; }, () => { sheet.Cells = before; sheet.RowCount = count; });
    }

    public void DeleteColumn(EditorSheet sheet, int col)
    {
        if (!IsCsv) throw new InvalidOperationException("Excel 数据为只读，请在 Excel 中编辑。");
        if (sheet.Columns.Count <= 1) throw new InvalidOperationException("请至少保留一列。");
        var beforeCells = sheet.Cells;
        var beforeColumns = sheet.Columns;
        var afterCells = new Dictionary<string, EditorCell>(StringComparer.OrdinalIgnoreCase);
        foreach (var (key, cell) in beforeCells)
        {
            var p = EditorSheet.Position(key);
            if (p.Column != col) afterCells[EditorSheet.Address(p.Row, p.Column > col ? p.Column - 1 : p.Column)] = cell;
        }
        string colName = beforeColumns[col].ColName ?? $"#{col + 1}";
        var afterColumns = beforeColumns.Where((_, i) => i != col).ToList();
        Execute($"删除列 {colName}", () => { sheet.Cells = afterCells; sheet.Columns = afterColumns; }, () => { sheet.Cells = beforeCells; sheet.Columns = beforeColumns; });
    }
}
