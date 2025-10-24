param()

Add-Type -AssemblyName System.Windows.Forms

# Check if snipping tool is already running
$existingProcess = Get-Process -Name "SnippingTool" -ErrorAction SilentlyContinue
$wasRunning = $existingProcess -ne $null

if ($wasRunning) {
    Write-Output "ALREADY_RUNNING"
} else {
    Write-Output "NOT_RUNNING"
    # Start snipping tool
    Start-Process "snippingtool.exe"
}

# Wait for snipping tool to be ready (check up to 5 seconds)
$maxAttempts = 50
$attempt = 0
$found = $false

while ($attempt -lt $maxAttempts -and -not $found) {
    $process = Get-Process -Name "SnippingTool" -ErrorAction SilentlyContinue
    if ($process) {
        $found = $true
        Write-Output "DETECTED"
    } else {
        Start-Sleep -Milliseconds 100
        $attempt++
    }
}

if ($found) {
    # Wait a bit longer for window to be ready
    Start-Sleep -Milliseconds 500

    # Send Win+Shift+S keyboard combination
    # Using Windows API to send keys more reliably
    Add-Type @"
        using System;
        using System.Runtime.InteropServices;
        public class KeyboardSend {
            [DllImport("user32.dll")]
            public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);

            public const int KEYEVENTF_KEYDOWN = 0x0000;
            public const int KEYEVENTF_KEYUP = 0x0002;
            public const int VK_LWIN = 0x5B;
            public const int VK_SHIFT = 0x10;
            public const int VK_S = 0x53;

            public static void SendWinShiftS() {
                // Press Win
                keybd_event(VK_LWIN, 0, KEYEVENTF_KEYDOWN, 0);
                System.Threading.Thread.Sleep(50);

                // Press Shift
                keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYDOWN, 0);
                System.Threading.Thread.Sleep(50);

                // Press S
                keybd_event(VK_S, 0, KEYEVENTF_KEYDOWN, 0);
                System.Threading.Thread.Sleep(50);

                // Release S
                keybd_event(VK_S, 0, KEYEVENTF_KEYUP, 0);
                System.Threading.Thread.Sleep(50);

                // Release Shift
                keybd_event(VK_SHIFT, 0, KEYEVENTF_KEYUP, 0);
                System.Threading.Thread.Sleep(50);

                // Release Win
                keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0);
            }
        }
"@

    [KeyboardSend]::SendWinShiftS()
    Write-Output "KEYS_SENT"
} else {
    Write-Output "TIMEOUT"
}
