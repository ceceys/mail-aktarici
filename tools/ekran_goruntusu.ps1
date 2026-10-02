# Uygulamayı açar, pencerenin görüntüsünü PNG olarak kaydeder, kapatır (README görseli için).
# Kullanım: pwsh tools\ekran_goruntusu.ps1 -Exe dist\MailAktarici.exe -Cikti docs\ekran.png
param([string]$Exe = 'dist\MailAktarici.exe', [string]$Cikti = 'docs\ekran.png', [int]$Bekle = 4)
Add-Type -AssemblyName System.Drawing
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class Win {
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int L, T, R, B; }
    [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h, out RECT r);
    [DllImport("user32.dll")] public static extern bool PrintWindow(IntPtr h, IntPtr hdc, uint flags);
    [DllImport("dwmapi.dll")] public static extern int DwmGetWindowAttribute(IntPtr h, int attr, out RECT r, int size);
}
'@
$p = Start-Process $Exe -PassThru
try {
    $deadline = (Get-Date).AddSeconds(20)
    do { Start-Sleep -Milliseconds 300; $p.Refresh() } while ($p.MainWindowHandle -eq 0 -and (Get-Date) -lt $deadline)
    Start-Sleep -Seconds $Bekle
    $h = $p.MainWindowHandle
    $r = New-Object Win+RECT
    [void][Win]::GetWindowRect($h, [ref]$r)
    $w = $r.R - $r.L; $hh = $r.B - $r.T
    $bmp = New-Object System.Drawing.Bitmap $w, $hh
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $hdc = $g.GetHdc()
    [void][Win]::PrintWindow($h, $hdc, 2)
    $g.ReleaseHdc($hdc); $g.Dispose()
    # Windows 11 gölge kenarını kırp
    $vis = New-Object Win+RECT
    if ([Win]::DwmGetWindowAttribute($h, 9, [ref]$vis, 16) -eq 0) {
        $crop = New-Object System.Drawing.Rectangle ($vis.L - $r.L), ($vis.T - $r.T), ($vis.R - $vis.L), ($vis.B - $vis.T)
        $c = $bmp.Clone($crop, $bmp.PixelFormat); $bmp.Dispose(); $bmp = $c
    }
    New-Item -ItemType Directory -Force (Split-Path -Parent $Cikti) | Out-Null
    $bmp.Save($Cikti, [System.Drawing.Imaging.ImageFormat]::Png); $bmp.Dispose()
    "Kaydedildi: $Cikti ($w x $hh)"
}
finally { if (-not $p.HasExited) { [void]$p.CloseMainWindow(); Start-Sleep 1; if (-not $p.HasExited) { $p.Kill() } } }
