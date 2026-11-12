using GridCraft.WinForms;

int failed = 0, passed = 0;
void Check(string what, bool ok, string? detail = null)
{
    if (ok) passed++; else failed++;
    Console.WriteLine($"  {(ok ? "OK  " : "FAIL")} {what}{(ok || detail == null ? "" : "  -> " + detail)}");
}

Console.WriteLine("Finding gridcraft.exe");
const string app = @"C:\Apps\Logbook";
string Resolve(string? path) => GridCraftControl.ResolveExePath(path, app);

Check("the default is taken from the app's folder", Resolve("gridcraft.exe") == @"C:\Apps\Logbook\gridcraft.exe", Resolve("gridcraft.exe"));
Check("whatever the current directory is",
    Directory.GetCurrentDirectory() != app && Resolve("gridcraft.exe").StartsWith(app, StringComparison.OrdinalIgnoreCase));
Check("a relative folder is under the app's folder too", Resolve(@"tools\gridcraft.exe") == @"C:\Apps\Logbook\tools\gridcraft.exe", Resolve(@"tools\gridcraft.exe"));
Check("and may climb out of it", Resolve(@"..\GridCraft\gridcraft.exe") == @"C:\Apps\GridCraft\gridcraft.exe", Resolve(@"..\GridCraft\gridcraft.exe"));
Check("a full path is kept as given", Resolve(@"D:\Tools\gridcraft.exe") == @"D:\Tools\gridcraft.exe");
Check("a blank path means gridcraft.exe in the app's folder", Resolve("  ") == @"C:\Apps\Logbook\gridcraft.exe" && Resolve(null) == @"C:\Apps\Logbook\gridcraft.exe");
Check("without a folder given, the running app's own folder",
    GridCraftControl.ResolveExePath("gridcraft.exe") == Path.Combine(AppContext.BaseDirectory, "gridcraft.exe"));

Console.WriteLine();
Console.WriteLine($"{passed} passed, {failed} failed.");
return failed == 0 ? 0 : 1;
