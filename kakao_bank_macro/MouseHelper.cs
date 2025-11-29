using System;
using System.Runtime.InteropServices;

public class MouseHelper
{
    [StructLayout(LayoutKind.Sequential)]
    struct INPUT
    {
        public int type;
        public MOUSEINPUT mi;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct MOUSEINPUT
    {
        public int dx;
        public int dy;
        public int mouseData;
        public int dwFlags;
        public int time;
        public IntPtr dwExtraInfo;
    }

    const int INPUT_MOUSE = 0;
    const int MOUSEEVENTF_MOVE = 0x0001;
    const int MOUSEEVENTF_ABSOLUTE = 0x8000;
    const int MOUSEEVENTF_LEFTDOWN = 0x0002;
    const int MOUSEEVENTF_LEFTUP = 0x0004;

    [DllImport("user32.dll", SetLastError = true)]
    static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

    // ★ 화면 좌표를 SendInput이 사용하는 0~65535 절대값으로 변환
    private static int ToAbsoluteX(int x)
    {
        return (int)(x * 65535 / GetSystemMetrics(0));
    }

    private static int ToAbsoluteY(int y)
    {
        return (int)(y * 65535 / GetSystemMetrics(1));
    }

    [DllImport("user32.dll")]
    static extern int GetSystemMetrics(int nIndex);

    // ★ 마우스 이동
    private static void MoveMouseAbsolute(int x, int y)
    {
        INPUT[] inputs = new INPUT[1];
        inputs[0].type = INPUT_MOUSE;
        inputs[0].mi.dx = ToAbsoluteX(x);
        inputs[0].mi.dy = ToAbsoluteY(y);
        inputs[0].mi.dwFlags = MOUSEEVENTF_MOVE | MOUSEEVENTF_ABSOLUTE;

        SendInput(1, inputs, Marshal.SizeOf(typeof(INPUT)));
    }

    // ★ 마우스 왼쪽 DOWN / UP
    private static void LeftDown()
    {
        INPUT[] inputs = new INPUT[1];
        inputs[0].type = INPUT_MOUSE;
        inputs[0].mi.dwFlags = MOUSEEVENTF_LEFTDOWN;
        SendInput(1, inputs, Marshal.SizeOf(typeof(INPUT)));
    }

    private static void LeftUp()
    {
        INPUT[] inputs = new INPUT[1];
        inputs[0].type = INPUT_MOUSE;
        inputs[0].mi.dwFlags = MOUSEEVENTF_LEFTUP;
        SendInput(1, inputs, Marshal.SizeOf(typeof(INPUT)));
    }

    // ★ 최종 드래그 함수
    public static void Drag(int x1, int y1, int x2, int y2, int x3, int y3, int x4, int y4)
    {
        // 1. 시작 지점 이동
        MoveMouseAbsolute(x1, y1);
        System.Threading.Thread.Sleep(1000);

        // 2. 누르기
        LeftDown();
        System.Threading.Thread.Sleep(1000);

        // 3. 드래그 이동
        MoveMouseAbsolute(x2, y2);
        System.Threading.Thread.Sleep(1000);

        MoveMouseAbsolute(x3, y3);
        System.Threading.Thread.Sleep(1000);

        MoveMouseAbsolute(x4, y4);
        System.Threading.Thread.Sleep(1000);

        // 4. 떼기
        LeftUp();
    }
}