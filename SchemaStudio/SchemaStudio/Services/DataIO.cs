using CsvHelper;
using ExcelDataReader;
using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Globalization;
using System.Text;

namespace SchemaStudio.Services
{
    internal static class DataIO
    {
        // —— 读 —— //
        public static List<string> ReadCsvHeader(string path)
        {
            using var reader = new StreamReader(path, DetectEncoding(path));
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            if (!csv.Read()) return new();
            var header = new List<string>();
            for (int i = 0; csv.TryGetField<string>(i, out var val); i++) header.Add((val ?? $"col_{i}").Trim());
            return header;
        }
        public static List<string[]> ReadCsvRows(string path, int expectedCols)
        {
            var list = new List<string[]>();
            using var reader = new StreamReader(path, DetectEncoding(path));
            using var csv = new CsvReader(reader, CultureInfo.InvariantCulture);
            if (!csv.Read()) return list;
            while (csv.Read())
            {
                var row = new string[expectedCols];
                for (int i = 0; i < expectedCols; i++) { csv.TryGetField<string>(i, out var v); row[i] = v?.Trim() ?? ""; }
                list.Add(row);
            }
            return list;
        }
        public static List<string> ReadXlsxHeader(string path, string sheet)
        {
            using var fs = File.OpenRead(path);
            using var r  = ExcelReaderFactory.CreateReader(fs);
            var ds = r.AsDataSet();
            var t  = ds.Tables.Cast<DataTable>().FirstOrDefault(x => string.Equals(x.TableName, sheet, StringComparison.OrdinalIgnoreCase)) ?? ds.Tables[0];
            if (t.Rows.Count == 0) return new();
            var row = t.Rows[0];
            var list = new List<string>();
            for (int i = 0; i < t.Columns.Count; i++) list.Add((row[i]?.ToString() ?? $"col_{i}").Trim());
            return list;
        }
        public static List<string[]> ReadXlsxRows(string path, string sheet, int expectedCols)
        {
            var list = new List<string[]>();
            using var fs = File.OpenRead(path);
            using var r  = ExcelReaderFactory.CreateReader(fs);
            var ds = r.AsDataSet();
            var t  = ds.Tables.Cast<DataTable>().FirstOrDefault(x => string.Equals(x.TableName, sheet, StringComparison.OrdinalIgnoreCase)) ?? ds.Tables[0];
            if (t.Rows.Count <= 1) return list;
            for (int rr = 1; rr < t.Rows.Count; rr++)
            {
                var dr = t.Rows[rr];
                var row = new string[expectedCols];
                for (int c = 0; c < expectedCols; c++) row[c] = dr[c]?.ToString()?.Trim() ?? "";
                list.Add(row);
            }
            return list;
        }

        // —— 写 —— //
        public static void WriteCsvHeader(string path, IEnumerable<string> headers)
        {
            using var sw = new StreamWriter(path, false, new UTF8Encoding(true));
            static string Esc(string s) =>
                (s.Contains('"') || s.Contains(',') || s.Contains('\n') || s.Contains('\r'))
                    ? "\"" + s.Replace("\"", "\"\"") + "\"" : s;
            sw.WriteLine(string.Join(",", headers.Select(Esc)));
        }

        // 每个 sheet 自己的列
        public static bool TryWriteXlsxMulti(string path, IEnumerable<(string sheet, string[] headers)> sheets, out string warn)
        {
            warn = "";
            using var wb = new XSSFWorkbook();
            foreach (var (sheet, headers) in sheets)
            {
                var ws = wb.CreateSheet(sheet);
                var r0 = ws.CreateRow(0);
                for (int c = 0; c < headers.Length; c++)
                    r0.CreateCell(c).SetCellValue(headers[c]);
            }
            using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
            wb.Write(fs);
            return true;
        }
        public static bool TryWriteXlsMulti(string path, IEnumerable<(string sheet, string[] headers)> sheets, out string warn)
        {
            warn = "";
            using var wb = new HSSFWorkbook();
            foreach (var (sheet, headers) in sheets)
            {
                var ws = wb.CreateSheet(sheet);
                var r0 = ws.CreateRow(0);
                for (int c = 0; c < headers.Length; c++)
                    r0.CreateCell(c).SetCellValue(headers[c]);
            }
            using var fs = new FileStream(path, FileMode.Create, FileAccess.Write);
            wb.Write(fs);
            return true;
        }

        // —— 杂项 —— //
        public static string GetSchemaJsonPath(string tablePath)
            => Path.Combine(Path.GetDirectoryName(tablePath)!, Path.GetFileName(tablePath) + ".schema.json");

        public static IEnumerable<string> GetAllDirs(string root)
        {
            yield return root;
            foreach (var d in Directory.GetDirectories(root))
                foreach (var x in GetAllDirs(d)) yield return x;
        }

        public static Encoding DetectEncoding(string filePath)
        {
            var bytes = File.ReadAllBytes(filePath);
            if (bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) return new UTF8Encoding(true);
            if (bytes.Length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) return Encoding.Unicode;
            if (bytes.Length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF) return Encoding.BigEndianUnicode;
            try
            {
                var strictUtf8 = Encoding.GetEncoding("UTF-8", new EncoderExceptionFallback(), new DecoderExceptionFallback());
                _ = strictUtf8.GetString(bytes);
                return new UTF8Encoding(false);
            }
            catch
            {
                return Encoding.GetEncoding(54936); // GB18030
            }
        }

        public static void SaveRegistry(string path, Models.PropertyRegistry reg)
        {
            var json = JsonConvert.SerializeObject(reg, Formatting.Indented,
                new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
            File.WriteAllText(path, json, new UTF8Encoding(false));
        }
        public static Models.PropertyRegistry LoadRegistry(string path)
        {
            if (!File.Exists(path)) return new Models.PropertyRegistry();
            var txt = File.ReadAllText(path);
            return JsonConvert.DeserializeObject<Models.PropertyRegistry>(txt) ?? new Models.PropertyRegistry();
        }
    }
}
