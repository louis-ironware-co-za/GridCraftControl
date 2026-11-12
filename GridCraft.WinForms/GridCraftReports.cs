using System.Data;
using System.Globalization;
using System.Text;
using System.Text.Json.Nodes;

namespace GridCraft.WinForms;

/// <summary>Options for <see cref="GridCraftReports.CreateReportAsync"/>.</summary>
public sealed class ReportOptions
{
    /// <summary>Table style of each sheet's data (Excel's built-in style names).</summary>
    public string TableStyle { get; set; } = "TableStyleMedium2";

    /// <summary>Number format of DateTime columns without a time of day.</summary>
    public string DateFormat { get; set; } = "yyyy-mm-dd";

    /// <summary>Number format of DateTime columns where some value has a time of day.</summary>
    public string DateTimeFormat { get; set; } = "yyyy-mm-dd hh:mm";

    /// <summary>Number format of decimal, double and float columns.</summary>
    public string DecimalFormat { get; set; } = "#,##0.00";

    /// <summary>Number format of integer columns ("0": no thousands separator, so IDs read as written).</summary>
    public string IntegerFormat { get; set; } = "0";

    /// <summary>Writes the report name as a title above each sheet's table.</summary>
    public bool Title { get; set; } = true;

    /// <summary>Adds a total row (sums of the numeric columns) to each table.</summary>
    public bool TotalRow { get; set; }

    /// <summary>When set, the report is saved there (.xlsx, .pdf, .csv … by extension).</summary>
    public string? SavePath { get; set; }
}

/// <summary>Reports from DataSets, and reading sheets back into DataTables.</summary>
public static class GridCraftReports
{
    private const int MaxRows = 1_048_576, MaxCols = 16_384;

    /// <summary>Rows per range.setValues call, so one request line stays a few MB.</summary>
    private const int ChunkRows = 2_000;

    /// <summary>
    /// Builds a new workbook from <paramref name="data"/> and shows it: one sheet per DataTable,
    /// named after the table, with the report name as a title, the columns as an Excel table
    /// (header row, filters, banding), number formats by column type, frozen headers and
    /// fitted column widths.
    /// </summary>
    public static async Task CreateReportAsync(this GridCraftClient c, DataSet data, string reportName, ReportOptions? options = null, CancellationToken ct = default)
    {
        ArgumentNullException.ThrowIfNull(data);
        var o = options ?? new ReportOptions();
        var tables = data.Tables.Cast<DataTable>().ToList();
        if (tables.Count == 0) tables.Add(new DataTable(string.IsNullOrWhiteSpace(reportName) ? "Report" : reportName));

        await c.ExecuteAsync("file.new", new { }, ct);
        var sheetNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var tableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 0; i < tables.Count; i++)
        {
            var sheet = UniqueName(SheetName(tables[i].TableName, i), sheetNames, 31);
            if (i == 0) await c.ExecuteAsync("sheet.rename", new { name = sheet }, ct);
            else await c.ExecuteAsync("home.insertSheet", new { name = sheet, before = i }, ct);
            var tableName = UniqueName(TableName(tables[i].TableName, i), tableNames, 255, identifier: true);
            await WriteTableAsync(c, tables[i], reportName, tableName, o, ct);
        }
        await c.ExecuteAsync("sheet.activate", new { sheet = 0 }, ct);
        await c.ExecuteAsync("selection.set", new { range = "A1" }, ct);
        if (!string.IsNullOrEmpty(o.SavePath)) await c.SaveAsync(o.SavePath, ct);
    }

    /// <summary>Fills the active sheet with one DataTable (see <see cref="CreateReportAsync"/>).</summary>
    private static async Task WriteTableAsync(GridCraftClient c, DataTable t, string reportName, string tableName, ReportOptions o, CancellationToken ct)
    {
        var cols = Math.Max(t.Columns.Count, 1);
        if (cols > MaxCols) throw new ArgumentException($"Table '{t.TableName}' has {t.Columns.Count} columns; a sheet holds {MaxCols}.");
        var headerRow = o.Title ? 3 : 1;
        if ((long)headerRow + Math.Max(t.Rows.Count, 1) > MaxRows)
            throw new ArgumentException($"Table '{t.TableName}' has {t.Rows.Count} rows; a sheet holds {MaxRows - headerRow}.");
        var last = ColumnLetters(cols - 1);

        // Header: column captions, unique (an Excel table needs unique header names).
        var headers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var header = t.Columns.Count == 0
            ? new object?[] { "(no columns)" }
            : t.Columns.Cast<DataColumn>().Select(dc => (object?)UniqueName(string.IsNullOrWhiteSpace(dc.Caption) ? dc.ColumnName : dc.Caption, headers, 255)).ToArray();
        await c.CallAsync("range.setValues", new { range = $"A{headerRow}", values = new[] { header }, parse = false }, ct);

        // Data, in chunks. Text that starts with '=' would be taken as a formula, so it is
        // entered afterwards as text ('…), never run.
        var formulaLike = new List<(string Cell, string Text)>();
        for (var start = 0; start < t.Rows.Count; start += ChunkRows)
        {
            var count = Math.Min(ChunkRows, t.Rows.Count - start);
            var block = new object?[count][];
            for (var r = 0; r < count; r++)
            {
                var row = t.Rows[start + r];
                var cells = new object?[t.Columns.Count];
                for (var col = 0; col < t.Columns.Count; col++)
                {
                    var v = ToCell(row[col]);
                    if (v is string s && s.StartsWith('='))
                    {
                        formulaLike.Add(($"{ColumnLetters(col)}{headerRow + 1 + start + r}", s));
                        v = null;
                    }
                    cells[col] = v;
                }
                block[r] = cells;
            }
            await c.CallAsync("range.setValues", new { range = $"A{headerRow + 1 + start}", values = block, parse = false }, ct);
        }
        foreach (var (cell, text) in formulaLike)
            await c.SetCellAsync(cell, "'" + text, ct);

        var lastRow = headerRow + Math.Max(t.Rows.Count, 1);
        for (var col = 0; col < t.Columns.Count; col++)
        {
            var format = NumberFormat(t, t.Columns[col], o);
            if (format is null || t.Rows.Count == 0) continue;
            var letters = ColumnLetters(col);
            await c.ExecuteAsync("home.numberFormat", new { range = $"{letters}{headerRow + 1}:{letters}{lastRow}", code = format }, ct);
        }

        await c.ExecuteAsync("insert.table", new { range = $"A{headerRow}:{last}{lastRow}", header = true, style = o.TableStyle, name = tableName }, ct);
        // Sums under decimal columns only: totals of IDs (integers), dates or text mean nothing.
        var summed = t.Columns.Cast<DataColumn>().Where(dc => IsNumeric(Nullable.GetUnderlyingType(dc.DataType) ?? dc.DataType) && !IsInteger(dc.DataType)).ToList();
        if (o.TotalRow && t.Rows.Count > 0 && summed.Count > 0)
        {
            await c.ExecuteAsync("table.totalRow", new { table = tableName, on = true }, ct);
            // Turning the row on puts a function under the last column; set every column explicitly.
            foreach (DataColumn dc in t.Columns)
                await c.ExecuteAsync("table.totalFunction", new { table = tableName, column = header[dc.Ordinal]!.ToString(), function = summed.Contains(dc) ? "sum" : "none" }, ct);
        }
        await c.ExecuteAsync("view.freezePanes", new { cell = $"A{headerRow + 1}" }, ct);
        // Fit the columns to the data before the title goes in, so a long title doesn't widen column A.
        await c.ExecuteAsync("home.autofitColumnWidth", new { cols = $"A:{last}" }, ct);
        if (o.Title)
        {
            await c.CallAsync("range.setValues", new { range = "A1", values = new[] { new object?[] { reportName } }, parse = false }, ct);
            await c.ExecuteAsync("home.bold", new { range = "A1", on = true }, ct);
            await c.ExecuteAsync("home.fontSize", new { range = "A1", size = 14 }, ct);
        }
    }

    /// <summary>
    /// Reads a sheet (default: the active one) or a range of it into a DataTable. With
    /// <paramref name="hasHeader"/> the first row names the columns. Column types are inferred:
    /// all numbers → double, all booleans → bool, otherwise string. Dates come back as numbers
    /// (Excel serial dates); use <see cref="DateTime.FromOADate"/> on columns you know hold dates.
    /// </summary>
    public static async Task<DataTable> ReadTableAsync(this GridCraftClient c, string? sheet = null, string? range = null, bool hasHeader = true, CancellationToken ct = default)
    {
        // Without a range, a sheet with an Excel table (as reports have) reads that table, not the
        // title above it, and leaves out its total row.
        if (range is null && hasHeader) range = await FirstTableRangeAsync(c, sheet, ct);
        var p = new JsonObject();
        if (sheet is not null) p["sheet"] = sheet;
        if (range is not null) p["range"] = range;
        var result = await c.CallAsync("sheet.read", p, ct);
        var rows = result?["values"]?.AsArray().Select(r => r?.AsArray().ToArray() ?? []).ToList() ?? [];
        var table = new DataTable(result?["sheet"]?.GetValue<string>() ?? sheet ?? "Sheet1");
        if (rows.Count == 0) return table;

        var width = rows.Max(r => r.Length);
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var dataRows = hasHeader ? rows.Skip(1).ToList() : rows;
        // Blank rows at the end (an empty table keeps one) are not data.
        while (dataRows.Count > 0 && dataRows[^1].All(v => v is null || (v is JsonValue jv && jv.TryGetValue<string>(out var s) && s.Length == 0)))
            dataRows.RemoveAt(dataRows.Count - 1);
        for (var col = 0; col < width; col++)
        {
            var name = hasHeader && col < rows[0].Length && rows[0][col] is { } h ? h.ToString() : "";
            if (string.IsNullOrWhiteSpace(name)) name = ColumnLetters(col);
            var cells = dataRows.Select(r => col < r.Length ? r[col] : null).Where(v => v is not null).ToList();
            var type = cells.Count > 0 && cells.All(IsNumber) ? typeof(double)
                : cells.Count > 0 && cells.All(IsBool) ? typeof(bool)
                : typeof(string);
            table.Columns.Add(UniqueName(name, names, 255), type);
        }
        foreach (var r in dataRows)
        {
            var values = new object[width];
            for (var col = 0; col < width; col++)
            {
                var v = col < r.Length ? r[col] : null;
                var type = table.Columns[col].DataType;
                values[col] = v is null ? DBNull.Value
                    : type == typeof(double) ? v.GetValue<double>()
                    : type == typeof(bool) ? v.GetValue<bool>()
                    : v is JsonValue jv && jv.TryGetValue<string>(out var s) ? s : v.ToJsonString();
            }
            table.Rows.Add(values);
        }
        return table;
    }

    /// <summary>The first Excel table's range on the sheet (default: active), without its total row; null if none.</summary>
    private static async Task<string?> FirstTableRangeAsync(GridCraftClient c, string? sheet, CancellationToken ct)
    {
        var doc = await c.InspectDocumentAsync(ct);
        var sheets = doc?["sheets"]?.AsArray();
        if (sheets is null) return null;
        var s = sheet is null
            ? sheets.ElementAtOrDefault(doc?["activeSheet"]?.GetValue<int>() ?? 0)
            : sheets.FirstOrDefault(x => string.Equals(x?["name"]?.GetValue<string>(), sheet, StringComparison.OrdinalIgnoreCase));
        var table = s?["tables"]?.AsArray().FirstOrDefault();
        var range = table?["range"]?.GetValue<string>();
        if (range is null) return null;
        if (table?["totals"]?.GetValue<bool>() != true) return range;
        var m = System.Text.RegularExpressions.Regex.Match(range, @"^(\$?[A-Z]+\$?\d+:\$?[A-Z]+\$?)(\d+)$");
        return m.Success && int.TryParse(m.Groups[2].Value, out var end) ? $"{m.Groups[1].Value}{end - 1}" : range;
    }

    /// <summary>Reads every sheet into a DataSet (one DataTable per sheet; see <see cref="ReadTableAsync"/>).</summary>
    public static async Task<DataSet> ReadDataSetAsync(this GridCraftClient c, bool hasHeader = true, CancellationToken ct = default)
    {
        var doc = await c.InspectDocumentAsync(ct);
        var set = new DataSet();
        foreach (var s in doc?["sheets"]?.AsArray() ?? [])
        {
            var name = s?["name"]?.GetValue<string>();
            if (name is null || s?["usedRange"] is null) continue;
            set.Tables.Add(await c.ReadTableAsync(name, null, hasHeader, ct));
        }
        return set;
    }

    private static object? ToCell(object? v) => v switch
    {
        null or DBNull => null,
        string s => s,
        bool b => b,
        DateTime d => d.ToOADate(),
        DateTimeOffset d => d.DateTime.ToOADate(),
        DateOnly d => d.ToDateTime(TimeOnly.MinValue).ToOADate(),
        TimeSpan ts => ts.TotalDays,
        TimeOnly t => t.ToTimeSpan().TotalDays,
        decimal m => (double)m,
        double x when double.IsFinite(x) => x,
        float x when float.IsFinite(x) => (double)x,
        double or float => null,
        byte or sbyte or short or ushort or int or uint or long or ulong => Convert.ToDouble(v, CultureInfo.InvariantCulture),
        byte[] => "(binary)",
        _ => Convert.ToString(v, CultureInfo.InvariantCulture),
    };

    private static string? NumberFormat(DataTable t, DataColumn dc, ReportOptions o)
    {
        var type = Nullable.GetUnderlyingType(dc.DataType) ?? dc.DataType;
        if (type == typeof(DateTime) || type == typeof(DateTimeOffset))
        {
            var hasTime = t.Rows.Cast<DataRow>().Any(r => r[dc] switch
            {
                DateTime d => d.TimeOfDay != TimeSpan.Zero,
                DateTimeOffset d => d.TimeOfDay != TimeSpan.Zero,
                _ => false,
            });
            return hasTime ? o.DateTimeFormat : o.DateFormat;
        }
        if (type == typeof(DateOnly)) return o.DateFormat;
        if (type == typeof(TimeSpan) || type == typeof(TimeOnly)) return "hh:mm:ss";
        if (IsInteger(type)) return o.IntegerFormat;
        if (IsNumeric(type)) return o.DecimalFormat;
        return null;
    }

    private static bool IsInteger(Type t) =>
        t == typeof(byte) || t == typeof(sbyte) || t == typeof(short) || t == typeof(ushort) ||
        t == typeof(int) || t == typeof(uint) || t == typeof(long) || t == typeof(ulong);

    private static bool IsNumeric(Type t) => IsInteger(t) || t == typeof(decimal) || t == typeof(double) || t == typeof(float);

    private static bool IsNumber(JsonNode? v) => v is JsonValue jv && jv.TryGetValue<double>(out _);

    private static bool IsBool(JsonNode? v) => v is JsonValue jv && jv.TryGetValue<bool>(out _);

    /// <summary>A, B, … Z, AA, … for a zero-based column index.</summary>
    public static string ColumnLetters(int index)
    {
        var sb = new StringBuilder();
        for (var n = index + 1; n > 0; n = (n - 1) / 26) sb.Insert(0, (char)('A' + (n - 1) % 26));
        return sb.ToString();
    }

    /// <summary>A valid sheet name: no []:*?/\, not empty, at most 31 characters, not "History".</summary>
    private static string SheetName(string name, int index)
    {
        var s = new string((name ?? "").Select(ch => "[]:*?/\\".IndexOf(ch) >= 0 || char.IsControl(ch) ? ' ' : ch).ToArray());
        s = System.Text.RegularExpressions.Regex.Replace(s, @"\s+", " ").Trim().Trim('\'');
        if (s.Length > 31) s = s[..31].TrimEnd();
        if (s.Length == 0 || s.Equals("History", StringComparison.OrdinalIgnoreCase)) s = $"Sheet{index + 1}";
        return s;
    }

    /// <summary>A valid table name: letters, digits, '_' and '.', starting with a letter or '_'.</summary>
    private static string TableName(string name, int index)
    {
        var s = new string((name ?? "").Select(ch => char.IsLetterOrDigit(ch) || ch is '_' or '.' ? ch : '_').ToArray());
        if (s.Length == 0 || !(char.IsLetter(s[0]) || s[0] == '_')) s = "Table_" + s;
        // Names that look like cell references (A1, R1C1) are not allowed.
        if (System.Text.RegularExpressions.Regex.IsMatch(s, @"^([A-Za-z]{1,3}\d+|[RrCc]\d*([RrCc]\d*)?)$")) s = "Table_" + s;
        return s.Length > 200 ? s[..200] : (s == "Table_" ? $"Table{index + 1}" : s);
    }

    /// <summary>
    /// <paramref name="name"/>, or with " (2)", " (3)" … (table names, which allow no spaces or
    /// brackets: "_2", "_3" …) so it differs from every name in <paramref name="used"/>.
    /// </summary>
    private static string UniqueName(string name, HashSet<string> used, int maxLength, bool identifier = false)
    {
        var candidate = name.Length > maxLength ? name[..maxLength] : name;
        for (var n = 2; !used.Add(candidate); n++)
        {
            var suffix = identifier ? $"_{n}" : $" ({n})";
            candidate = (name.Length + suffix.Length > maxLength ? name[..(maxLength - suffix.Length)] : name) + suffix;
        }
        return candidate;
    }
}
