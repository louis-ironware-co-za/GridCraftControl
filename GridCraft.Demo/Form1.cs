using GridCraft.WinForms;

namespace GridCraft.Demo;

/// <summary>Demo: a GridCraft spreadsheet docked in a WinForms form, driven from C#.</summary>
public partial class Form1 : Form
{
    private readonly GridCraftControl _grid = new() { Dock = DockStyle.Fill };
    // Empty = act on the cell(s) selected in GridCraft, like a ribbon button. Type a cell to target it instead.
    private readonly TextBox _cell = new() { Width = 80, PlaceholderText = "selection" };
    private readonly TextBox _input = new() { Text = "=SUM(1,2,3)", Width = 220 };
    private readonly Label _status = new() { AutoSize = true, Padding = new Padding(8, 6, 0, 0) };
    private readonly TextBox _output = new()
    {
        Dock = DockStyle.Bottom, Height = 110, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
        Font = new Font(FontFamily.GenericMonospace, 9f),
    };

    public Form1()
    {
        Text = "GridCraft in WinForms";
        Width = 1400;
        Height = 900;

        // Default: the release build next to this solution (../gridcraft-main/target/release).
        _grid.ExePath = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            @"..\..\..\..\..\gridcraft-main\target\release\gridcraft.exe"));
        _grid.Sample = "sales";
        _grid.GridCraftExited += (_, _) => _status.Text = "GridCraft exited.";

        var bar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 36, Padding = new Padding(4), WrapContents = false };
        bar.Controls.Add(new Label { Text = "Cell", AutoSize = true, Padding = new Padding(0, 6, 0, 0) });
        bar.Controls.Add(_cell);
        bar.Controls.Add(_input);
        bar.Controls.Add(Button("Set", async c =>
        {
            var cell = await ActiveCell(c);
            return $"{cell}: " + ((await c.SetCellAsync(cell, _input.Text))?.ToJsonString() ?? "null");
        }));
        bar.Controls.Add(Button("Get", async c =>
        {
            var cell = await ActiveCell(c);
            var r = await c.GetCellAsync(cell);
            return $"{cell}: value={r?["value"]?.ToJsonString()}  input={r?["input"]?.ToJsonString()}\r\n{r?.ToJsonString()}";
        }));
        bar.Controls.Add(Button("Bold", async c =>
        {
            var range = await Range(c);
            return $"{range}: " + ((await c.ExecuteAsync("home.bold", new { range }))?.ToJsonString() ?? "null");
        }));
        bar.Controls.Add(Button("Sample report", async c =>
        {
            await c.CreateReportAsync(SampleData(), "Monthly orders", new ReportOptions { TotalRow = true });
            return "report created";
        }));
        bar.Controls.Add(Button("Read sheet", async c =>
        {
            var t = await c.ReadTableAsync();
            var cols = string.Join(", ", t.Columns.Cast<System.Data.DataColumn>().Select(dc => $"{dc.ColumnName} ({dc.DataType.Name})"));
            var first = t.Rows.Cast<System.Data.DataRow>().Take(5).Select(r => "  " + string.Join(" | ", r.ItemArray));
            return $"'{t.TableName}': {t.Rows.Count} rows\r\n{cols}\r\n{string.Join("\r\n", first)}";
        }));
        bar.Controls.Add(Button("Open…", async c =>
        {
            using var dlg = new OpenFileDialog { Filter = "Workbooks|*.xlsx;*.xlsm;*.csv|All files|*.*" };
            if (dlg.ShowDialog(this) != DialogResult.OK) return "";
            return (await c.OpenAsync(dlg.FileName))?.ToJsonString() ?? "null";
        }));
        bar.Controls.Add(Button("Save as…", async c =>
        {
            using var dlg = new SaveFileDialog { Filter = "Excel workbook|*.xlsx", FileName = "Book1.xlsx" };
            if (dlg.ShowDialog(this) != DialogResult.OK) return "";
            return (await c.SaveAsync(dlg.FileName))?.ToJsonString() ?? "null";
        }));
        bar.Controls.Add(_status);

        Controls.Add(_grid);
        Controls.Add(_output);
        Controls.Add(bar);

        Load += async (_, _) =>
        {
            try
            {
                _status.Text = "Starting GridCraft…";
                await _grid.StartAsync();
                _status.Text = "Connected.";
            }
            catch (Exception ex)
            {
                _status.Text = "Start failed";
                _output.Text = ex.ToString();
            }
        };
    }

    /// <summary>A small DataSet standing in for one of your report queries.</summary>
    private static System.Data.DataSet SampleData()
    {
        var ds = new System.Data.DataSet();
        var orders = ds.Tables.Add("Orders");
        orders.Columns.Add("OrderID", typeof(int));
        orders.Columns.Add("Customer", typeof(string));
        orders.Columns.Add("Ordered", typeof(DateTime));
        orders.Columns.Add("Amount", typeof(decimal));
        orders.Columns.Add("Paid", typeof(bool));
        var names = new[] { "Acme (Pty) Ltd", "Bateman", "Northwind", "Contoso Build" };
        for (var i = 0; i < 40; i++)
            orders.Rows.Add(1001 + i, names[i % names.Length], new DateTime(2026, 3, 1).AddDays(i), 250m + i * 137.5m, i % 3 != 0);
        var visits = ds.Tables.Add("Site visits");
        visits.Columns.Add("When", typeof(DateTime));
        visits.Columns.Add("Site", typeof(string));
        visits.Columns.Add("Hours", typeof(double));
        visits.Rows.Add(new DateTime(2026, 3, 3, 9, 30, 0), "Block A", 2.5);
        visits.Rows.Add(new DateTime(2026, 3, 10, 14, 0, 0), "Block C", 1.25);
        return ds;
    }

    /// <summary>The cell typed in the Cell box, else GridCraft's active cell.</summary>
    private async Task<string> ActiveCell(GridCraftClient c) =>
        string.IsNullOrWhiteSpace(_cell.Text) ? (await c.GetSelectionAsync()).ActiveCell : _cell.Text.Trim();

    /// <summary>The range typed in the Cell box, else GridCraft's whole selection.</summary>
    private async Task<string> Range(GridCraftClient c) =>
        string.IsNullOrWhiteSpace(_cell.Text) ? (await c.GetSelectionAsync()).Selection : _cell.Text.Trim();

    private Button Button(string text, Func<GridCraftClient, Task<string>> action)
    {
        var b = new Button { Text = text, AutoSize = true };
        b.Click += async (_, _) =>
        {
            if (_grid.Client is not { } c) { _status.Text = "Not connected."; return; }
            _status.Text = text + "…";
            try
            {
                var result = await action(c);
                if (result.Length == 0) { _status.Text = ""; return; } // dialog cancelled
                _status.Text = text + " OK";
                _output.Text = $"{text} {result}";
            }
            catch (Exception ex)
            {
                _status.Text = text + " failed";
                _output.Text = $"{text}: {ex.GetType().Name}: {ex.Message}";
            }
        };
        return b;
    }
}
