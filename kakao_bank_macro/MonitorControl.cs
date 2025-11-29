using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

namespace kakao_bank_macro
{
    internal class MonitorControl
    {
        const int HWND_BROADCAST = 0xffff;
        const int WM_SYSCOMMAND = 0x0112;
        const int SC_MONITORPOWER = 0xF170;

        [DllImport("user32.dll")]
        static extern bool SetCursorPos(int X, int Y);

        [DllImport("user32.dll")]
        static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, IntPtr lParam);

        // 모니터 끄기
        public static void TurnOffMonitor()
        {
            SendMessage((IntPtr)HWND_BROADCAST, WM_SYSCOMMAND,
                (IntPtr)SC_MONITORPOWER, (IntPtr)2);   // 2 = OFF
        }

        // 모니터 켜기
        public static void TurnOnMonitor()
        {
            SendMessage((IntPtr)HWND_BROADCAST, WM_SYSCOMMAND,
                (IntPtr)SC_MONITORPOWER, (IntPtr)(-1));  // -1 = ON
        }
    }
}
