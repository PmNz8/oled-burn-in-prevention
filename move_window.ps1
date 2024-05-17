Add-Type @"
using System;
using System.Runtime.InteropServices;

public class Win32 {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

    public struct RECT {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }
}
"@

function Get-ForegroundWindowRect {
    $hwnd = [Win32]::GetForegroundWindow()
    if ($hwnd -eq [IntPtr]::Zero) {
        return $null
    }

    $rect = New-Object Win32+RECT
    if (-not [Win32]::GetWindowRect($hwnd, [ref]$rect)) {
        return $null
    }

    return @{ hwnd = $hwnd; rect = $rect }
}

function Move-Window {
    param (
        [IntPtr]$hwnd,
        [int]$x,
        [int]$y,
        [int]$width,
        [int]$height
    )
    [Win32]::MoveWindow($hwnd, $x, $y, $width, $height, $true) | Out-Null
}

$deltaX = 1
$deltaY = 1
$screenWidth = 3440
$screenHeight = 1440

while ($true) {
    $windowInfo = Get-ForegroundWindowRect
    if ($null -eq $windowInfo) {
        Start-Sleep -Seconds 1
        continue
    }

    $hwnd = $windowInfo.hwnd
    $rect = $windowInfo.rect

    $windowWidth = $rect.Right - $rect.Left
    $windowHeight = $rect.Bottom - $rect.Top
    if ($windowWidth -ge $screenWidth -and $windowHeight -ge $screenHeight) {
        Start-Sleep -Seconds 30
        continue
    }

    $newX = $rect.Left + $deltaX
    $newY = $rect.Top + $deltaY

    if ($newX -lt 0 -or $newX + ($rect.Right - $rect.Left) -gt $screenWidth) {
        $deltaX = -$deltaX
        $newX = $rect.Left + $deltaX
    }

    if ($newY -lt 0 -or $newY + ($rect.Bottom - $rect.Top) -gt $screenHeight) {
        $deltaY = -$deltaY
        $newY = $rect.Top + $deltaY
    }

    Move-Window -hwnd $hwnd -x $newX -y $newY -width ($rect.Right - $rect.Left) -height ($rect.Bottom - $rect.Top)

    Start-Sleep -Seconds 1
}