using System.ComponentModel;
using System.Diagnostics;
using System.Net;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Security.Cryptography;

namespace GridCraft.WinForms;

/// <summary>
/// Hosts a GridCraft window inside a WinForms container. It starts <c>gridcraft.exe</c> with a
/// control port and a private token, docks the app's window into this control, and exposes
/// <see cref="Client"/> for driving it. Call <see cref="StartAsync"/> (e.g. from Form.Load).
/// </summary>
public class GridCraftControl : UserControl
{
    private Process? _process;
    private IntPtr _child;

    /// <summary>
    /// Path to gridcraft.exe. A relative path (the default, "gridcraft.exe") is taken from the
    /// application's own folder, not the current directory: an app started from a shortcut whose
    /// "Start in" is elsewhere still finds the gridcraft.exe shipped beside it.
    /// </summary>
    [Category("GridCraft"), DefaultValue("gridcraft.exe")]
    public string ExePath { get; set; } = "gridcraft.exe";

    /// <summary>
    /// The full path <see cref="ExePath"/> stands for: as given when rooted, otherwise under
    /// <paramref name="appFolder"/> (the application's folder, <see cref="AppContext.BaseDirectory"/>).
    /// </summary>
    public static string ResolveExePath(string? exePath, string? appFolder = null)
    {
        var path = string.IsNullOrWhiteSpace(exePath) ? "gridcraft.exe" : exePath.Trim();
        return Path.GetFullPath(Path.IsPathFullyQualified(path) ? path : Path.Combine(appFolder ?? AppContext.BaseDirectory, path));
    }

    /// <summary>Optional sample workbook to open at start: "sales", "budget" or "grades".</summary>
    [Category("GridCraft"), DefaultValue(null)]
    public string? Sample { get; set; }

    /// <summary>Optional workbook file to open at start.</summary>
    [Category("GridCraft"), DefaultValue(null)]
    public string? FilePath { get; set; }

    /// <summary>
    /// Starts GridCraft with <c>--embedded</c>: no title bar of its own, and every workbook opened
    /// stays in this control (switch with View › Switch Windows) instead of opening a new window.
    /// Needs a gridcraft.exe built with that option; an older one would try to open "--embedded" as a file.
    /// </summary>
    [Category("GridCraft"), DefaultValue(true)]
    public bool Embedded { get; set; } = true;

    private bool _showRibbon = true;

    /// <summary>
    /// GridCraft's own ribbon. Off, the host's commands stand in for it: GridCraft starts with
    /// <c>--no-ribbon</c>, and changing this while it runs takes effect at once (control channel
    /// <c>ui.set</c>). Keyboard shortcuts, the formula bar and the print page stay either way.
    /// Needs a gridcraft.exe built with that option.
    /// </summary>
    [Category("GridCraft"), DefaultValue(true)]
    public bool ShowRibbon
    {
        get => _showRibbon;
        set
        {
            if (_showRibbon == value) return;
            _showRibbon = value;
            if (Client != null) _ = Client.CallAsync("ui.set", new { ribbon = value });
        }
    }

    private bool? _dark;

    /// <summary>
    /// GridCraft's dark theme: true dark, false light, null (the default) as GridCraft chooses
    /// (it follows Windows). A host with its own theme sets this to match it. Set before start,
    /// it is sent as soon as GridCraft connects; set while it runs, it takes effect at once
    /// (control channel <c>ui.set {dark}</c>).
    /// </summary>
    [Category("GridCraft"), DefaultValue(null)]
    public bool? Dark
    {
        get => _dark;
        set
        {
            if (_dark == value) return;
            _dark = value;
            if (Client != null && value is bool dark) _ = Client.CallAsync("ui.set", new { dark });
        }
    }

    /// <summary>Opens GridCraft's print page: preview, printer and page settings, and its Print button.</summary>
    public Task ShowPrintPageAsync() =>
        Client?.ExecuteAsync("ui.backstage", new { page = "print" }) ?? Task.CompletedTask;

    /// <summary>How long to wait for GridCraft's window and control channel.</summary>
    [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public TimeSpan StartTimeout { get; set; } = TimeSpan.FromSeconds(30);

    /// <summary>The control-channel client; available after <see cref="StartAsync"/> completes.</summary>
    [Browsable(false), DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
    public GridCraftClient? Client { get; private set; }

    /// <summary>True while the GridCraft process is running.</summary>
    [Browsable(false)]
    public bool IsRunning => _process is { HasExited: false };

    /// <summary>Raised when the GridCraft process exits (for example the user closed the workbook).</summary>
    public event EventHandler? GridCraftExited;

    public GridCraftControl()
    {
        BackColor = SystemColors.AppWorkspace;
    }

    /// <summary>Starts GridCraft, docks its window here, and connects <see cref="Client"/>.</summary>
    public async Task StartAsync(CancellationToken ct = default)
    {
        if (DesignMode) return;
        if (IsRunning) throw new InvalidOperationException("GridCraft is already running in this control.");
        var exe = ResolveExePath(ExePath);
        if (!File.Exists(exe)) throw new FileNotFoundException($"gridcraft.exe not found at {exe}.", exe);

        var port = FreeLoopbackPort();
        var token = Convert.ToHexString(RandomNumberGenerator.GetBytes(32));

        var psi = new ProcessStartInfo(exe) { UseShellExecute = false, WorkingDirectory = Path.GetDirectoryName(exe) };
        psi.ArgumentList.Add("--control");
        psi.ArgumentList.Add(port.ToString());
        if (Embedded) psi.ArgumentList.Add("--embedded");
        if (!ShowRibbon) psi.ArgumentList.Add("--no-ribbon");
        if (!string.IsNullOrEmpty(Sample)) { psi.ArgumentList.Add("--sample"); psi.ArgumentList.Add(Sample); }
        if (!string.IsNullOrEmpty(FilePath)) psi.ArgumentList.Add(FilePath);
        // Passing the token through the environment keeps it out of the process list, and GridCraft then writes no token file.
        psi.Environment["GRIDCRAFT_CONTROL_TOKEN"] = token;

        _process = Process.Start(psi) ?? throw new InvalidOperationException("Could not start gridcraft.exe.");
        _process.EnableRaisingEvents = true;
        _process.Exited += (_, _) => BeginInvokeSafe(() => { _child = IntPtr.Zero; GridCraftExited?.Invoke(this, EventArgs.Empty); });

        _child = await WaitForMainWindowAsync(_process, StartTimeout, ct);
        EmbedWindow(_child);
        Client = await GridCraftClient.ConnectAsync(port, token, StartTimeout, ct);
        if (Dark is bool dark) await Client.CallAsync("ui.set", new { dark }, ct);
    }

    /// <summary>
    /// Asks GridCraft to quit, then kills it if it has not exited after <paramref name="wait"/>.
    /// Unsaved changes are discarded (programmatic quit never asks), so save first if needed.
    /// </summary>
    public void Stop(TimeSpan? wait = null)
    {
        var p = _process;
        if (p is null) return;
        try
        {
            if (!p.HasExited)
            {
                try { Client?.CallAsync("app.quit").Wait(TimeSpan.FromSeconds(2)); } catch { /* fall through to kill */ }
                if (!p.WaitForExit(wait ?? TimeSpan.FromSeconds(3))) p.Kill(entireProcessTree: true);
            }
        }
        catch (InvalidOperationException) { /* already gone */ }
        Client?.Dispose();
        Client = null;
        p.Dispose();
        _process = null;
        _child = IntPtr.Zero;
    }

    // -------------------------------------------------------------------------------------
    // BeginUpdate / EndUpdate
    // -------------------------------------------------------------------------------------

    private int _updateDepth;
    private PictureBox? _still;
    private bool _hadFocus;

    /// <summary>True between <see cref="BeginUpdate"/> and the matching <see cref="EndUpdateAsync"/>.</summary>
    [Browsable(false)]
    public bool IsUpdating => _updateDepth > 0;

    /// <summary>
    /// Stops the sheet showing changes until <see cref="EndUpdateAsync"/>, as DevExpress's
    /// BeginUpdate does: a batch of commands (adding rows, formatting a sheet) is seen once,
    /// finished, instead of step by step. GridCraft runs in its own process and can't pause its
    /// drawing, so this shows a still picture of the sheet in its place and moves the real window
    /// out of sight, where it keeps running the commands. Calls nest; only the outermost one counts.
    /// </summary>
    public void BeginUpdate()
    {
        if (_updateDepth++ > 0 || _child == IntPtr.Zero || ClientSize.Width <= 0 || ClientSize.Height <= 0) return;

        var picture = new Bitmap(ClientSize.Width, ClientSize.Height);
        using (var g = Graphics.FromImage(picture))
        {
            // PrintWindow asks the window for its own content (works for GPU-drawn windows and when
            // something overlaps it); the screen is the fallback.
            var hdc = g.GetHdc();
            bool printed;
            try { printed = Native.PrintWindow(_child, hdc, Native.PW_RENDERFULLCONTENT); }
            finally { g.ReleaseHdc(hdc); }
            if (!printed) g.CopyFromScreen(PointToScreen(Point.Empty), Point.Empty, picture.Size);
        }
        _still = new PictureBox { Dock = DockStyle.Fill, Image = picture, SizeMode = PictureBoxSizeMode.Normal };
        Controls.Add(_still);
        _still.BringToFront();
        _hadFocus = ContainsFocus;
        // Moved out of sight rather than hidden: a hidden window runs GridCraft's commands at a
        // background pace (a month took 6 s instead of 1). See OutOfSight.
        Native.MoveWindow(_child, OutOfSight, 0, ClientSize.Width, ClientSize.Height, false);
    }

    /// <summary>
    /// Where the window waits during an update: all but its first column past this control's
    /// right edge, under the still picture. Fully outside, Windows counts it as covered and
    /// GridCraft slows down; one column inside keeps it drawing at full pace.
    /// </summary>
    private int OutOfSight => Math.Max(0, ClientSize.Width - 1);

    /// <summary>
    /// Shows the sheet again once GridCraft has run every command sent before this call. Ends
    /// the outermost <see cref="BeginUpdate"/>; an inner one only counts down.
    /// </summary>
    public async Task EndUpdateAsync()
    {
        if (_updateDepth == 0 || --_updateDepth > 0) return;
        try
        {
            // One more round trip: every command sent before it has been carried out by then.
            if (Client != null) await Client.CallAsync("ui.inspect");
        }
        catch (Exception ex) when (ex is IOException or GridCraftException or ObjectDisposedException)
        {
            // GridCraft gone or busy: show whatever there is rather than a stale picture.
        }
        if (_child != IntPtr.Zero)
        {
            Native.MoveWindow(_child, 0, 0, ClientSize.Width, ClientSize.Height, true);
            // A frame or two for GridCraft to draw the finished sheet before the picture goes.
            await Task.Delay(80);
            if (_hadFocus) Native.SetFocus(_child);
        }
        if (_still != null)
        {
            Controls.Remove(_still);
            _still.Image?.Dispose();
            _still.Dispose();
            _still = null;
        }
    }

    /// <summary>Runs <paramref name="work"/> between BeginUpdate and EndUpdateAsync, ending the update even when it fails.</summary>
    public async Task UpdateAsync(Func<Task> work)
    {
        BeginUpdate();
        try { await work(); }
        finally { await EndUpdateAsync(); }
    }

    protected override void OnResize(EventArgs e)
    {
        base.OnResize(e);
        // While updating, the window stays out of sight (BeginUpdate) at the new size.
        if (_child != IntPtr.Zero)
            Native.MoveWindow(_child, _still != null ? OutOfSight : 0, 0, ClientSize.Width, ClientSize.Height, _still == null);
    }

    protected override void OnGotFocus(EventArgs e)
    {
        base.OnGotFocus(e);
        if (_child != IntPtr.Zero) Native.SetFocus(_child);
    }

    /// <summary>
    /// A docked window only gets keyboard focus if it takes it when clicked, and GridCraft's window
    /// (winit) doesn't: it expects to be top-level. Windows tells the parent about mouse-downs in a
    /// child (WM_PARENTNOTIFY), so on a click inside GridCraft we move keyboard focus to it.
    /// SetFocus works across processes here because SetParent attached the two input queues.
    /// </summary>
    protected override void WndProc(ref Message m)
    {
        if (m.Msg == Native.WM_PARENTNOTIFY && _child != IntPtr.Zero)
        {
            var evt = (int)(m.WParam.ToInt64() & 0xFFFF);
            if (evt is Native.WM_LBUTTONDOWN or Native.WM_RBUTTONDOWN or Native.WM_MBUTTONDOWN)
            {
                if (!ContainsFocus) Focus(); // let WinForms know focus moved here (OnGotFocus focuses the child)
                Native.SetFocus(_child);
            }
        }
        base.WndProc(ref m);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) Stop();
        base.Dispose(disposing);
    }

    private void EmbedWindow(IntPtr hwnd)
    {
        var style = Native.GetWindowLongPtr(hwnd, Native.GWL_STYLE).ToInt64();
        style &= ~(Native.WS_CAPTION | Native.WS_THICKFRAME | Native.WS_SYSMENU | Native.WS_MINIMIZEBOX | Native.WS_MAXIMIZEBOX | Native.WS_POPUP);
        style |= Native.WS_CHILD | Native.WS_VISIBLE;
        Native.SetWindowLongPtr(hwnd, Native.GWL_STYLE, new IntPtr(style));
        Native.SetParent(hwnd, Handle);
        Native.SetWindowPos(hwnd, IntPtr.Zero, 0, 0, ClientSize.Width, ClientSize.Height,
            Native.SWP_NOZORDER | Native.SWP_FRAMECHANGED | Native.SWP_SHOWWINDOW);
    }

    private static async Task<IntPtr> WaitForMainWindowAsync(Process p, TimeSpan timeout, CancellationToken ct)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            ct.ThrowIfCancellationRequested();
            if (p.HasExited) throw new InvalidOperationException($"gridcraft.exe exited with code {p.ExitCode} before showing its window.");
            var hwnd = FindAppWindow((uint)p.Id);
            if (hwnd != IntPtr.Zero) return hwnd;
            await Task.Delay(100, ct);
        }
        throw new TimeoutException("GridCraft did not show a window in time.");
    }

    /// <summary>
    /// The app's visible, titled top-level window. Not <c>Process.MainWindowHandle</c>: winit first
    /// creates a hidden "Winit Thread Event Target" helper window, which that property can return.
    /// </summary>
    private static IntPtr FindAppWindow(uint pid)
    {
        var found = IntPtr.Zero;
        Native.EnumWindows((h, _) =>
        {
            Native.GetWindowThreadProcessId(h, out var owner);
            if (owner != pid || !Native.IsWindowVisible(h) || Native.GetWindowTextLength(h) == 0) return true;
            var cls = new System.Text.StringBuilder(256);
            Native.GetClassName(h, cls, cls.Capacity);
            if (cls.ToString() is "Winit Thread Event Target" or "IME" or "MSCTFIME UI") return true;
            found = h;
            return false;
        }, IntPtr.Zero);
        return found;
    }

    private static int FreeLoopbackPort()
    {
        var l = new TcpListener(IPAddress.Loopback, 0);
        l.Start();
        var port = ((IPEndPoint)l.LocalEndpoint).Port;
        l.Stop();
        return port;
    }

    private void BeginInvokeSafe(Action a)
    {
        if (IsDisposed || !IsHandleCreated) return;
        try { BeginInvoke(a); } catch (InvalidOperationException) { }
    }

    private static class Native
    {
        public const int GWL_STYLE = -16;
        public const int WM_PARENTNOTIFY = 0x0210, WM_LBUTTONDOWN = 0x0201, WM_RBUTTONDOWN = 0x0204, WM_MBUTTONDOWN = 0x0207;
        public const long WS_CHILD = 0x40000000L, WS_VISIBLE = 0x10000000L, WS_POPUP = 0x80000000L;
        public const long WS_CAPTION = 0x00C00000L, WS_THICKFRAME = 0x00040000L, WS_SYSMENU = 0x00080000L;
        public const long WS_MINIMIZEBOX = 0x00020000L, WS_MAXIMIZEBOX = 0x00010000L;
        public const uint SWP_NOZORDER = 0x0004, SWP_FRAMECHANGED = 0x0020, SWP_SHOWWINDOW = 0x0040;
        public const uint PW_RENDERFULLCONTENT = 2;

        [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr hwnd, IntPtr hdc, uint flags);

        [DllImport("user32.dll", SetLastError = true)] public static extern IntPtr SetParent(IntPtr child, IntPtr parent);
        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtrW")] public static extern IntPtr GetWindowLongPtr(IntPtr hwnd, int index);
        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")] public static extern IntPtr SetWindowLongPtr(IntPtr hwnd, int index, IntPtr value);
        [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr hwnd, IntPtr after, int x, int y, int cx, int cy, uint flags);
        [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr hwnd, int x, int y, int w, int h, bool repaint);
        [DllImport("user32.dll")] public static extern IntPtr SetFocus(IntPtr hwnd);

        public delegate bool EnumWindowsProc(IntPtr hwnd, IntPtr lParam);
        [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc proc, IntPtr lParam);
        [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint pid);
        [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hwnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowTextLength(IntPtr hwnd);
        [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr hwnd, System.Text.StringBuilder name, int max);
    }
}
