using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace GridCraft.WinForms;

/// <summary>Error returned by GridCraft for a request (<c>{"ok":false,"error":"…"}</c>).</summary>
public sealed class GridCraftException(string message) : Exception(message);

/// <summary>
/// Client for GridCraft's JSON-lines control channel (see docs/control-protocol.md in the
/// GridCraft repository). One request is in flight at a time; calls are serialized.
/// </summary>
public sealed class GridCraftClient : IDisposable
{
    private readonly TcpClient _tcp;
    private readonly StreamReader _reader;
    private readonly StreamWriter _writer;
    private readonly SemaphoreSlim _lock = new(1, 1);
    private long _nextId = 1;

    private GridCraftClient(TcpClient tcp)
    {
        _tcp = tcp;
        var stream = tcp.GetStream();
        _reader = new StreamReader(stream, new UTF8Encoding(false));
        _writer = new StreamWriter(stream, new UTF8Encoding(false)) { NewLine = "\n", AutoFlush = true };
    }

    /// <summary>Connects to <c>127.0.0.1:port</c> and authenticates, retrying until <paramref name="timeout"/>.</summary>
    public static async Task<GridCraftClient> ConnectAsync(int port, string token, TimeSpan timeout, CancellationToken ct = default)
    {
        var deadline = DateTime.UtcNow + timeout;
        Exception? last = null;
        while (DateTime.UtcNow < deadline)
        {
            ct.ThrowIfCancellationRequested();
            var tcp = new TcpClient { NoDelay = true };
            try
            {
                await tcp.ConnectAsync("127.0.0.1", port, ct);
                var client = new GridCraftClient(tcp);
                await client.AuthenticateAsync(token, ct);
                return client;
            }
            catch (Exception e) when (e is SocketException or IOException or GridCraftException)
            {
                last = e;
                tcp.Dispose();
                await Task.Delay(200, ct);
            }
        }
        throw new TimeoutException($"Could not connect to GridCraft on port {port}.", last);
    }

    private async Task AuthenticateAsync(string token, CancellationToken ct)
    {
        var line = JsonSerializer.Serialize(new { id = 0, method = "auth", @params = new { token } });
        await _writer.WriteLineAsync(line.AsMemory(), ct);
        var reply = await _reader.ReadLineAsync(ct) ?? throw new IOException("Connection closed during authentication.");
        var node = JsonNode.Parse(reply);
        if (node?["ok"]?.GetValue<bool>() != true)
            throw new GridCraftException(node?["error"]?.GetValue<string>() ?? "authentication failed");
    }

    /// <summary>Calls any control method (e.g. <c>ui.inspect</c>, <c>sheet.read</c>) and returns its <c>result</c>.</summary>
    public async Task<JsonNode?> CallAsync(string method, object? parameters = null, CancellationToken ct = default)
    {
        await _lock.WaitAsync(ct);
        try
        {
            var id = _nextId++;
            var request = new JsonObject
            {
                ["id"] = id,
                ["method"] = method,
                ["params"] = parameters is null ? new JsonObject() : JsonSerializer.SerializeToNode(parameters),
            };
            await _writer.WriteLineAsync(request.ToJsonString().AsMemory(), ct);
            while (true)
            {
                var line = await _reader.ReadLineAsync(ct) ?? throw new IOException("GridCraft closed the connection.");
                if (string.IsNullOrWhiteSpace(line)) continue;
                var reply = JsonNode.Parse(line);
                if (reply?["id"]?.GetValue<long>() != id) continue; // a late reply to an earlier, cancelled call
                if (reply["ok"]?.GetValue<bool>() == true) return reply["result"];
                throw new GridCraftException(reply["error"]?.GetValue<string>() ?? "unknown error");
            }
        }
        finally
        {
            _lock.Release();
        }
    }

    /// <summary>Runs an engine command, e.g. <c>ExecuteAsync("home.bold", new { range = "A1:C1" })</c>.</summary>
    public Task<JsonNode?> ExecuteAsync(string command, object? parameters = null, CancellationToken ct = default) =>
        CallAsync("engine.execute", new { command, @params = parameters ?? new { } }, ct);

    /// <summary>Types <paramref name="input"/> into a cell, as the user would (a value or a formula starting with '=').</summary>
    public Task<JsonNode?> SetCellAsync(string cell, string input, CancellationToken ct = default) =>
        ExecuteAsync("cell.set", new { cell, input }, ct);

    /// <summary>One cell: value, formula, displayed text, style, merge, comment, link.</summary>
    public Task<JsonNode?> GetCellAsync(string cell, CancellationToken ct = default) =>
        CallAsync("cell.get", new { cell }, ct);

    /// <summary>Values (or formulas / displayed text) of a range such as "A1:D10".</summary>
    public Task<JsonNode?> ReadRangeAsync(string range, bool formulas = false, bool formatted = false, CancellationToken ct = default) =>
        CallAsync("sheet.read", new { range, formulas, formatted }, ct);

    /// <summary>The user's current selection in the active workbook, e.g. ("B5:C7", "B5").</summary>
    public async Task<(string Selection, string ActiveCell)> GetSelectionAsync(CancellationToken ct = default)
    {
        var doc = await InspectDocumentAsync(ct);
        return (doc?["selection"]?.GetValue<string>() ?? "A1", doc?["activeCell"]?.GetValue<string>() ?? "A1");
    }

    /// <summary>Workbook summary: sheets, used ranges, tables, charts, names, selection.</summary>
    public Task<JsonNode?> InspectDocumentAsync(CancellationToken ct = default) => CallAsync("document.inspect", null, ct);

    /// <summary>Opens a workbook (.xlsx and the other formats GridCraft reads).</summary>
    public Task<JsonNode?> OpenAsync(string path, string? password = null, CancellationToken ct = default) =>
        CallAsync("app.open", password is null ? new { path } : new { path, password }, ct);

    /// <summary>Saves the active workbook; with a path it saves to that file.</summary>
    public Task<JsonNode?> SaveAsync(string? path = null, CancellationToken ct = default) =>
        CallAsync("app.save", path is null ? new { } : new { path }, ct);

    /// <summary>Lists every command (id, label, ribbon path, shortcut, params doc), optionally filtered.</summary>
    public Task<JsonNode?> CommandsAsync(string? search = null, CancellationToken ct = default) =>
        CallAsync("engine.commands", search is null ? null : new { search }, ct);

    public void Dispose()
    {
        _reader.Dispose();
        _writer.Dispose();
        _tcp.Dispose();
        _lock.Dispose();
    }
}
