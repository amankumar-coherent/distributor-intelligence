# Tile every automation Chromium window into a grid.
#
# All eleven browsers launch at the same default position and size, so
# they land exactly on top of one another. Only the topmost is visible,
# which reads as "only 6 are open" even though all eleven exist.
#
# Usage:  powershell -ExecutionPolicy Bypass -File tile_browsers.ps1

Add-Type -AssemblyName System.Windows.Forms

Add-Type @"
using System;
using System.Runtime.InteropServices;
public class WinTile {
  [DllImport("user32.dll")] public static extern bool MoveWindow(IntPtr h, int x, int y, int w, int t, bool repaint);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
}
"@

$screen = [System.Windows.Forms.Screen]::PrimaryScreen
$W = $screen.WorkingArea.Width
$H = $screen.WorkingArea.Height

# Only OUR automation browsers, identified by their profile directory,
# so the user's own Chrome windows are left alone.
$procs = Get-CimInstance Win32_Process -Filter "Name='chrome.exe'" |
  Where-Object { $_.CommandLine -match 'drone_profile_|ai_mode_chrome_profile' } |
  Select-Object -ExpandProperty ProcessId

$windows = @(Get-Process chrome -ErrorAction SilentlyContinue |
  Where-Object { $_.MainWindowHandle -ne 0 -and $procs -contains $_.Id })

$n = $windows.Count
if ($n -eq 0) {
  Write-Output "no automation windows found"
  exit
}

# 4 columns wide; rows follow from the count (11 -> 4x3).
$cols = 4
$rows = [math]::Ceiling($n / $cols)
$w = [int]($W / $cols)
$h = [int]($H / $rows)

$i = 0
foreach ($win in $windows) {
  $x = ($i % $cols) * $w
  $y = [int]([math]::Floor($i / $cols)) * $h
  [WinTile]::ShowWindow($win.MainWindowHandle, 9) | Out-Null   # SW_RESTORE
  [WinTile]::MoveWindow($win.MainWindowHandle, $x, $y, $w, $h, $true) | Out-Null
  $i++
}

Write-Output ("tiled {0} automation windows in a {1}x{2} grid" -f $n, $cols, $rows)
