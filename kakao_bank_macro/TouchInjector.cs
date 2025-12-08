using System;
using System.Runtime.InteropServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;

public static class TouchInjector
{
    // ─── constants ─────────────────────────────
    const int TOUCH_FEEDBACK_DEFAULT = 0x1;

    const int PT_TOUCH = 2;

    const int POINTER_FLAG_NONE = 0x00000000;
    const int POINTER_FLAG_DOWN = 0x00010000;
    const int POINTER_FLAG_UPDATE = 0x00020000;
    const int POINTER_FLAG_UP = 0x00040000;
    const int POINTER_FLAG_INRANGE = 0x00000002;
    const int POINTER_FLAG_INCONTACT = 0x00000004;

    [Flags]
    enum TOUCH_MASK : uint
    {
        NONE = 0x00000000,
        CONTACTAREA = 0x00000001,
        ORIENTATION = 0x00000002,
        PRESSURE = 0x00000004
    }

    enum TOUCH_FLAGS : uint
    {
        NONE = 0
    }

    enum POINTER_BUTTON_CHANGE_TYPE : uint
    {
        NONE = 0
    }

    [StructLayout(LayoutKind.Sequential)]
    struct POINT
    {
        public int X;
        public int Y;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct RECT
    {
        public int left;
        public int top;
        public int right;
        public int bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct POINTER_INFO
    {
        public uint pointerType;
        public uint pointerId;
        public uint frameId;
        public int pointerFlags;
        public IntPtr sourceDevice;
        public IntPtr hwndTarget;
        public POINT ptPixelLocation;
        public POINT ptHimetricLocation;
        public POINT ptPixelLocationRaw;
        public POINT ptHimetricLocationRaw;
        public uint dwTime;
        public uint historyCount;
        public int InputData;
        public uint dwKeyStates;
        public ulong PerformanceCount;
        public POINTER_BUTTON_CHANGE_TYPE ButtonChangeType;
    }

    [StructLayout(LayoutKind.Sequential)]
    struct POINTER_TOUCH_INFO
    {
        public POINTER_INFO pointerInfo;
        public TOUCH_FLAGS touchFlags;
        public TOUCH_MASK touchMask;
        public RECT rcContact;
        public RECT rcContactRaw;
        public uint orientation;
        public uint pressure;
    }

    [DllImport("user32.dll")]
    static extern bool InitializeTouchInjection(uint maxCount, uint dwMode);

    [DllImport("user32.dll")]
    static extern bool InjectTouchInput(uint count, [In] POINTER_TOUCH_INFO[] contacts);

    static bool _initialized = false;

    public static void EnsureInitialized()
    {
        if (_initialized) return;

        // 한 번만 초기화
        if (!InitializeTouchInjection(10, TOUCH_FEEDBACK_DEFAULT))
        {
            throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        }
        _initialized = true;
    }

    private static POINTER_TOUCH_INFO CreateContact(uint id, int x, int y, int flags)
    {
        var contact = new POINTER_TOUCH_INFO();

        contact.pointerInfo.pointerType = PT_TOUCH;
        contact.pointerInfo.pointerId = id;
        contact.pointerInfo.pointerFlags = flags;
        contact.pointerInfo.ptPixelLocation.X = x;
        contact.pointerInfo.ptPixelLocation.Y = y;

        // 터치가 닿은 영역(손가락 크기) - 너무 작으면 인식이 애매해져서 약간 크게
        int radius = 5;
        contact.rcContact.left = x - radius;
        contact.rcContact.top = y - radius;
        contact.rcContact.right = x + radius;
        contact.rcContact.bottom = y + radius;

        contact.touchFlags = TOUCH_FLAGS.NONE;
        contact.touchMask = TOUCH_MASK.CONTACTAREA | TOUCH_MASK.ORIENTATION | TOUCH_MASK.PRESSURE;

        contact.orientation = 90;   // 대충 90도
        contact.pressure = 32000; // 0~65535 사이 아무 값 (0이면 무시될 수 있음)

        return contact;
    }

    /// <summary>
    /// 화면 좌표 기준 단일 손가락 드래그 제스처
    /// </summary>
    public static void TouchDrag((int x, int y)[] points, int steps = 30, int deplay = 5)
    {
        EnsureInitialized();

        if (points == null || points.Length < 2)
            throw new ArgumentException("최소 두 개 이상의 경로 좌표가 필요합니다.");

        uint id = 1;

        // 1️⃣ DOWN — 첫 번째 좌표에서 터치 시작
        var first = points[0];
        var contact = CreateContact(
            id,
            first.x,
            first.y,
            POINTER_FLAG_DOWN | POINTER_FLAG_INRANGE | POINTER_FLAG_INCONTACT
        );
        InjectTouchInput(1, new[] { contact });
        System.Threading.Thread.Sleep(30);


        // 2️⃣ UPDATE — 각 구간을 30 step 으로 나눠서 이동
        for (int segment = 0; segment < points.Length - 1; segment++)
        {
            var start = points[segment];
            var end = points[segment + 1];

            for (int i = 1; i <= steps; i++)
            {
                int x = start.x + (end.x - start.x) * i / steps;
                int y = start.y + (end.y - start.y) * i / steps;

                contact = CreateContact(
                    id,
                    x,
                    y,
                    POINTER_FLAG_UPDATE | POINTER_FLAG_INRANGE | POINTER_FLAG_INCONTACT
                );

                InjectTouchInput(1, new[] { contact });
                System.Threading.Thread.Sleep(deplay);
            }
        }


        // 3️⃣ UP — 마지막 좌표에서 터치 끝
        var last = points[^1];
        contact = CreateContact(
            id,
            last.x,
            last.y,
            POINTER_FLAG_UP
        );
        InjectTouchInput(1, new[] { contact });
    }

    public static void TouchClickWithColor(int x, int y, Color targetColor, CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            if (IsColorMatch(x, y, targetColor))
                break;

            System.Threading.Thread.Sleep(100);
        }
        

        EnsureInitialized();

        uint id = 1;

        var contact = CreateContact(
            id,
            x,
            y,
            POINTER_FLAG_DOWN | POINTER_FLAG_INRANGE | POINTER_FLAG_INCONTACT
        );

        InjectTouchInput(1, new[] { contact });
        System.Threading.Thread.Sleep(30); // 눌림 유지 감지 안정성을 위해 50ms 유지

        // 2️⃣ 터치 UP
        contact = CreateContact(
            id,
            x,
            y,
            POINTER_FLAG_UP
        );

        InjectTouchInput(1, new[] { contact });

        Thread.Sleep(200);
    }

    public static void TouchClick(int x, int y)
    {
        EnsureInitialized();

        uint id = 1;

        var contact = CreateContact(
            id,
            x,
            y,
            POINTER_FLAG_DOWN | POINTER_FLAG_INRANGE | POINTER_FLAG_INCONTACT
        );

        InjectTouchInput(1, new[] { contact });
        System.Threading.Thread.Sleep(30); // 눌림 유지 감지 안정성을 위해 50ms 유지

        // 2️⃣ 터치 UP
        contact = CreateContact(
            id,
            x,
            y,
            POINTER_FLAG_UP
        );

        InjectTouchInput(1, new[] { contact });

        Thread.Sleep(200);
    }

    public static bool IsColorMatch(int x, int y, Color targetColor, int tolerance = 7)
    {
        using (Bitmap bmp = new Bitmap(1, 1))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(x, y, 0, 0, new Size(1, 1));
            }

            Color c = bmp.GetPixel(0, 0);

            return Math.Abs(c.R - targetColor.R) <= tolerance &&
               Math.Abs(c.G - targetColor.G) <= tolerance &&
               Math.Abs(c.B - targetColor.B) <= tolerance;
        }
    }

    public static Color getColor(int x, int y)
    {
        using (Bitmap bmp = new Bitmap(1, 1))
        {
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(x, y, 0, 0, new Size(1, 1));
            }

            Color c = bmp.GetPixel(0, 0);

            return c;
        }
    }
}