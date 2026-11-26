# Run by the installer (build/installer.nsh) before it touches the install folder.
# Closes every process running from that folder, then checks no file in it is locked.
# Exit 0 = safe to install. Exit 1 = still blocked; the blockers are in $env:EA_OUT.
$ErrorActionPreference = 'Continue'
$dir = $env:EA_DIR.TrimEnd('\') + '\'
$exe = $env:EA_EXE
$self = [int]$env:EA_PID
$log = Join-Path $env:TEMP 'event-automator-install.log'
function Log($m) { Add-Content -Path $log -Value ("{0:HH:mm:ss.fff} {1}" -f (Get-Date), $m) }
Log "---- check $dir (installer pid $self)"

# A custom install folder gets "\Event Automator" appended, so the old copy (which the old
# uninstaller is about to empty) may sit elsewhere: take it from the uninstall entry.
if (-not (Test-Path -LiteralPath (Join-Path $dir $exe))) {
  $old = Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*' -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like "$([IO.Path]::GetFileNameWithoutExtension($exe))*" -and $_.UninstallString } | Select-Object -First 1
  if ($old) { $dir = [IO.Path]::GetDirectoryName(($old.UninstallString -split '"')[1]) + '\'; Log "old install at $dir" }
}
# Only trust a folder that really holds the app, never a drive root.
if (-not (Test-Path -LiteralPath (Join-Path $dir $exe)) -or $dir.Length -le 4) { Log 'no app here, nothing to close'; exit 0 }
function Get-Ours {
  Get-CimInstance Win32_Process | Where-Object {
    $_.ProcessId -ne $self -and ($_.Name -ieq $exe -or ($_.ExecutablePath -and ([EaRm]::LongPath($_.ExecutablePath)).ToLower().StartsWith($dirL)))
  }
}

Add-Type -TypeDefinition @'
using System; using System.Runtime.InteropServices; using System.Collections.Generic;
public static class EaRm {
  [StructLayout(LayoutKind.Sequential)] struct UNIQUE_PROCESS { public int pid; public System.Runtime.InteropServices.ComTypes.FILETIME start; }
  [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)] struct INFO {
    public UNIQUE_PROCESS p; [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)] public string app;
    [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)] public string svc; public int type; public int status; public int session; public bool restartable; }
  [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)] static extern int RmStartSession(out uint h, int f, string key);
  [DllImport("rstrtmgr.dll")] static extern int RmEndSession(uint h);
  [DllImport("rstrtmgr.dll", CharSet = CharSet.Unicode)] static extern int RmRegisterResources(uint h, uint nf, string[] files, uint na, UNIQUE_PROCESS[] a, uint ns, string[] s);
  [DllImport("rstrtmgr.dll")] static extern int RmGetList(uint h, out uint need, ref uint n, [In, Out] INFO[] info, ref uint reasons);
  public static List<string> Holders(string[] files) {
    var r = new List<string>(); uint h; if (RmStartSession(out h, 0, Guid.NewGuid().ToString()) != 0) return r;
    try {
      if (RmRegisterResources(h, (uint)files.Length, files, 0, null, 0, null) != 0) return r;
      uint need = 0, n = 0, why = 0; RmGetList(h, out need, ref n, null, ref why);
      if (need == 0) return r; var info = new INFO[need]; n = need;
      if (RmGetList(h, out need, ref n, info, ref why) == 0) for (int i = 0; i < n; i++) r.Add(info[i].p.pid + "|" + info[i].app);
    } finally { RmEndSession(h); }
    return r;
  }
  [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] static extern uint GetLongPathName(string s, System.Text.StringBuilder b, uint n);
  public static string LongPath(string s) { var b = new System.Text.StringBuilder(1024); return GetLongPathName(s, b, 1024) > 0 ? b.ToString() : s; }
}
'@ -ErrorAction SilentlyContinue
# Process paths are long names; $INSTDIR may arrive in 8.3 form (NAVEEN~1).
$dirL = ([EaRm]::LongPath($dir)).ToLower()

function Get-Locked {
  $locked = @()
  # Only files a running program keeps open; opening all ~10k makes antivirus scan each (20 s+).
  $held = '.exe', '.dll', '.pyd', '.asar', '.node', '.pak', '.dat', '.bin'
  foreach ($f in [IO.Directory]::EnumerateFiles($dir, '*', 'AllDirectories')) {
    if ($held -notcontains [IO.Path]::GetExtension($f).ToLower()) { continue }
    # Sharing violation = someone holds it. Read-only files throw UnauthorizedAccess: not a lock.
    try { ([IO.File]::Open($f, 'Open', 'ReadWrite', 'None')).Dispose() } catch [IO.IOException] { $locked += $f } catch {}
  }
  $locked
}

# On update the old uninstaller moves every file to %TEMP%\nsXXXX.tmp\old-install\<path>
# and aborts ("cannot be closed") if one is over 260 chars. It deletes them all anyway, so
# delete the ones it can't move first. \\?\ lifts the limit for us.
function Remove-Overlong {
  $room = 259 - ([IO.Path]::GetTempPath().Length + 'nsx0000.tmp\old-install\'.Length) - 5
  $lp = '\\?\' + $dir
  foreach ($f in [IO.Directory]::EnumerateFiles($lp, '*', 'AllDirectories')) {
    if ($f.Length - $lp.Length -gt $room) {
      try { [IO.File]::Delete($f); Log "removed long path: $($f.Substring($lp.Length))" } catch { Log "could not remove $f : $_" }
    }
  }
}

$deadline = (Get-Date).AddSeconds(25)
$start = Get-Date
do {
  $procs = @(Get-Ours)
  if ($procs.Count) {
    # Give it ~3 s to exit by itself (it may be quitting for the update), then force it.
    if (((Get-Date) - $start).TotalSeconds -gt 3) {
      foreach ($p in $procs) {
        Log "kill $($p.ProcessId) $($p.Name) $($p.ExecutablePath)"
        & taskkill.exe /F /T /PID $p.ProcessId 2>&1 | ForEach-Object { Log "  $_" }
      }
    }
    Start-Sleep -Milliseconds 500; continue
  }
  $locked = @(Get-Locked)
  if (-not $locked.Count) { Remove-Overlong; Log 'clear'; exit 0 }
  Log "locked: $($locked.Count) file(s), e.g. $($locked[0])"
  Start-Sleep -Milliseconds 1000
} while ((Get-Date) -lt $deadline)

# Still blocked: name what holds the files, so the person knows what to close.
$names = @()
foreach ($p in @(Get-Ours)) { $names += "$($p.Name) (pid $($p.ProcessId), maybe run as administrator)" }
$locked = @(Get-Locked)
if ($locked.Count) {
  $h = @()
  try { $h = [EaRm]::Holders([string[]]($locked | Select-Object -First 50)) } catch { Log "RM failed: $_" }
  foreach ($x in $h) { $pid_, $app = $x -split '\|', 2; $names += "$app (pid $pid_)" }
  if (-not $h.Count) { $names += "an unknown program (antivirus?) is holding $([IO.Path]::GetFileName($locked[0]))" }
  $locked | Select-Object -First 20 | ForEach-Object { Log "  locked $_" }
}
$names = $names | Select-Object -Unique
if (-not $names) { Remove-Overlong; Log 'clear at the end'; exit 0 }
Log "blocked by: $($names -join '; ')"
Set-Content -Path $env:EA_OUT -Value ($names -join ', ') -Encoding Ascii
exit 1
