param([string]$windowTitle, [string]$textToSend)

Add-Type -Language CSharp @"
using System;
using System.Text;
using System.Runtime.InteropServices;

public class CopyDataSender {
    [StructLayout(LayoutKind.Sequential)]
    public struct COPYDATASTRUCT {
        public IntPtr dwData;
        public int cbData;
        public IntPtr lpData;
    }

    [DllImport("user32.dll", SetLastError=true)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("user32.dll", SetLastError=true)]
    public static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, ref COPYDATASTRUCT lParam);

    public static void SendString(string windowTitle, string textToSend) {
        const int WM_COPYDATA = 0x004A;
        IntPtr hwnd = FindWindow(null, windowTitle);
        if (hwnd == IntPtr.Zero) {
            Console.WriteLine("ETCIM20 window not found");
            return;
        }

        byte[] sarr = Encoding.Default.GetBytes(textToSend + "\0");
        IntPtr lpData = Marshal.AllocHGlobal(sarr.Length);
        Marshal.Copy(sarr, 0, lpData, sarr.Length);

        COPYDATASTRUCT cds = new COPYDATASTRUCT();
        cds.dwData = IntPtr.Zero;
        cds.cbData = sarr.Length;
        cds.lpData = lpData;

        SendMessage(hwnd, WM_COPYDATA, IntPtr.Zero, ref cds);
        Marshal.FreeHGlobal(lpData);

        Console.WriteLine("Message sent to ETCIM20: " + textToSend);
    }
}
"@

[CopyDataSender]::SendString($windowTitle, $textToSend)
