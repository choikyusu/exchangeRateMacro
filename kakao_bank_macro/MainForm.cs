using Microsoft.Web.WebView2.Core;
using System.Diagnostics;
using System.Runtime.InteropServices;
using YamlDotNet.Core.Tokens;

namespace kakao_bank_macro
{
    public partial class MainForm : Form
    {
        private int errorCount = 0;

        private CancellationTokenSource ctsMain;

        // t, t2 스레드 핸들 저장
        private Thread threadMain;
        private Thread threadWeb;

        int width;
        [DllImport("user32.dll", SetLastError = true)]
        static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter,
         int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);


        const uint SWP_NOZORDER = 0x0004;
        const uint SWP_NOMOVE = 0x0002;
        const uint SWP_SHOWWINDOW = 0x0040;

        const int EM_REPLACESEL = 0x00C2;

        static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);

        const uint SWP_NOSIZE = 0x0001;
        const uint WM_KEYDOWN = 0x0100;
        const uint WM_KEYUP = 0x0101;
        const int VK_ENTER = 0x0D; // 엔터 키 코드


        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter,
            string lpszClass, string lpszWindow);

        private bool isRunning = false;

        private bool needInitApp = false;


        public MainForm()
        {

            InitializeComponent();

        }

        private void startButton_Click(object sender, EventArgs e)
        {

            timerLabel.Text = DateTime.Now.ToString("HH시 mm분 ss초");


            width = Screen.PrimaryScreen.Bounds.Width;

            IntPtr hWnd = FindWindow(null, "Samsung Flow");
            //IntPtr hWnd = FindWindow(null, "Galaxy S10");
            bool result = SetWindowPos(
            hWnd,
            HWND_TOPMOST,   // 항상 위로
            0, 0, 0, 0,     // 위치/크기 유지
            SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);

            needInitApp = checkBoxInitApp.Checked;


            if (ctsMain == null)
            {
                isRunning = true;
                ctsMain = new CancellationTokenSource();

                threadMain = new Thread(() => RunAutomationLoopAsync(ctsMain.Token).Wait())
                {
                    IsBackground = true
                };
                threadMain.SetApartmentState(ApartmentState.STA);
                threadMain.Start();

                threadWeb = new Thread(() => RunAutomationWebLoopAsync(ctsMain.Token).Wait())
                {
                    IsBackground = true
                };
                threadWeb.SetApartmentState(ApartmentState.STA);
                threadWeb.Start();

                // 감시자는 기존처럼 실행 (독립 실행)
                Thread t3 = new Thread(async () => await RunAutomationCheckErrorAsync())
                {
                    IsBackground = true
                };
                t3.SetApartmentState(ApartmentState.STA);
                t3.Start();
            }
        }

        private async Task RunAutomationLoopAsync(CancellationToken token)
        {
            IntPtr hWnd = FindWindow(null, "Samsung Flow");
            // IntPtr hWnd = FindWindow(null, "Galaxy S10");
            this.Invoke((Delegate)(() =>
            {
                SetWindowPos(hWnd, IntPtr.Zero, width - 408, 0, 408, 900, SWP_NOZORDER | SWP_SHOWWINDOW);
            }));

            if (needInitApp)
                await initPhoneAsync(token); // 토큰 전달

            while (!token.IsCancellationRequested)
            {
                try
                {
                    if (token.IsCancellationRequested) break;

                    var sw = Stopwatch.StartNew(); // 시간 측정 시작
                    DateTime now = DateTime.Now;
                    if (DateTime.Now.Hour >= 7 && Properties.Settings.Default.startMorning != DateTime.Now.ToString("yyyy-MM-dd"))
                    {
                        Properties.Settings.Default.startMorning = DateTime.Now.ToString("yyyy-MM-dd");
                        CBHighTextbox.Text = "";
                        CBLowTextbox.Text = "";
                    }

                    initHome(token); // 토큰 전달
                    if (token.IsCancellationRequested) break;

                    //if (DateTime.Now.Hour >= 9 && DateTime.Now.Minute >= 5 || (DateTime.Now.Hour >= 10 && DateTime.Now.Hour < 16))
                    //    getRateFromKakaoPay(token);

                    getRateFromSwitchwon(token); // 토큰 전달
                    if (token.IsCancellationRequested) break;

                    getRateFromKakaoBank(token); // 토큰 전달
                    if (token.IsCancellationRequested) break;

                    getRateFromToss(token); // 토큰 전달
                    if (token.IsCancellationRequested) break;

                    this.Invoke((Delegate)(() =>
                    {
                        // sendKakaotalkMessage("최규수");
                        if (now.DayOfWeek >= DayOfWeek.Monday &&
                           now.Hour >= 7 || now.DayOfWeek >= DayOfWeek.Tuesday)
                        {
                            sendKakaotalkMessage("NEW환도박방");
                        }
                        else
                        {
                            Logger.Instance.Log("아직 시간 안됨");
                        }
                    }));


                    saveValue();

                    sw.Stop();
                    int elapsedMs = (int)sw.ElapsedMilliseconds;   // 걸린 시간(ms)
                    Logger.Instance.Log("루프 1회 도는데 " + elapsedMs + "ms 소요");
                    //int targetMs = 20000; // 20초 주기
                    //int remain = targetMs - elapsedMs;

                    //if (remain > 0)
                    //{
                    //    // WaitOne이 true를 반환하면 취소 신호를 받은 것 -> break
                    //    if (token.WaitHandle.WaitOne(remain)) break;
                    //}

                }
                catch (Exception ex)
                {
                    if (token.IsCancellationRequested) break;
                    var msg = ex.InnerException?.ToString() ?? ex.ToString();
                    Task.Run(() => MessageBox.Show(msg)); // UI 스레드 강요 없음
                }

            }
            Logger.Instance.Log("RunAutomationLoopAsync 루프 종료됨");
        }

        // CancellationToken 추가 및 Sleep -> WaitOne 변경
        private async Task initPhoneAsync(CancellationToken token)
        {
            Logger.Instance.Log("초기화: 폰 초기화시작");

            // Task.Delay도 취소 토큰 지원
            try { await Task.Delay(3000, token); } catch (TaskCanceledException) { return; }

            if (token.IsCancellationRequested) return;

            Logger.Instance.Log("초기화: 탭전환 클릭");
            TouchInjector.TouchClick(1597, 851);
            if (token.WaitHandle.WaitOne(3000)) return;

            Logger.Instance.Log("초기화: 모두 닫기 클릭");
            TouchInjector.TouchClick(1717, 689);
            if (token.WaitHandle.WaitOne(1000)) return;

            Logger.Instance.Log("초기화: 탭전환 클릭");
            TouchInjector.TouchClick(1597, 851);
            if (token.WaitHandle.WaitOne(3000)) return;

            Logger.Instance.Log("초기화: 자물쇠 클릭");
            TouchInjector.TouchClick(1807, 585);
            if (token.WaitHandle.WaitOne(2000)) return;

            Logger.Instance.Log("초기화: 모두 닫기 클릭");
            TouchInjector.TouchClick(1717, 689);
            if (token.WaitHandle.WaitOne(2000)) return;

            initSwitchwonApp(token);
            if (token.IsCancellationRequested) return;

            initKakaoBankApp(token);
            if (token.IsCancellationRequested) return;

            initTossApp(token);
            if (token.IsCancellationRequested) return;


            Logger.Instance.Log("초기화: 탭전환 클릭");
            TouchInjector.TouchClick(1597, 851);
            if (token.WaitHandle.WaitOne(3000)) return;

            Logger.Instance.Log("초기화: 토스앱 클릭");
            TouchInjector.TouchClick(1719, 193);
            if (token.WaitHandle.WaitOne(3000)) return;

            Logger.Instance.Log("초기화: 토스앱 고정");
            TouchInjector.TouchClick(1733, 333);
            if (token.WaitHandle.WaitOne(3000)) return;

            Logger.Instance.Log("초기화: 홈 클릭");
            TouchInjector.TouchClick(1713, 853);
            if (token.WaitHandle.WaitOne(3000)) return;


            Logger.Instance.Log("초기화: 세팅 끝");
        }

        private void initKakaoBankApp(CancellationToken token)
        {
            if (token.WaitHandle.WaitOne(1000)) return;
            Logger.Instance.Log("초기화: 카카오뱅크 시작" + TouchInjector.getColor(1585, 778).ToString());
            TouchInjector.TouchClickWithColor(1585, 778, Color.FromArgb(254, 227, 0), token);

            if (TouchInjector.IsColorMatch(1863, 783, Color.FromArgb(14, 168, 255)))
            {
                Logger.Instance.Log("초기화: 카카오뱅크 시작안됨" + TouchInjector.getColor(1863, 783).ToString());
                while (true)
                {
                    if (token.IsCancellationRequested) return; // 취소 확인

                    if (TouchInjector.IsColorMatch(1863, 783, Color.FromArgb(17, 170, 255)))
                    {
                        Logger.Instance.Log("초기화: 카카오뱅크 시작 다시 클릭" + TouchInjector.getColor(1863, 783).ToString());
                        TouchInjector.TouchClick(1585, 778);
                    }
                    else break;
                    if (token.WaitHandle.WaitOne(1000)) return;
                }
            }


            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (TouchInjector.IsColorMatch(1890, 867, Color.FromArgb(0, 0, 0)))
                {
                    Logger.Instance.Log("초기화: 카카오뱅크 검은색화면 표시" + TouchInjector.getColor(1890, 867).ToString());
                    break;
                }
                if (token.WaitHandle.WaitOne(100)) return;
            }

            Logger.Instance.Log("초기화: 카카오뱅크 터치시작");
            TouchInjector.TouchDrag(new (int x, int y)[]
                        {
                            (1622, 543),
                            (1822, 543),
                            (1617, 722),
                            (1822, 722)
                        }, 7, 20);


            Logger.Instance.Log("초기화: 카카오뱅크 환율화면 진입 시도");

            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (TouchInjector.IsColorMatch(1890, 867, Color.FromArgb(0, 0, 0)) && TouchInjector.IsColorMatch(1763, 449, Color.FromArgb(0, 0, 0)))
                {
                    Logger.Instance.Log("초기화: 카카오뱅크 검은색화면 표시" + TouchInjector.getColor(1890, 867).ToString());

                    Logger.Instance.Log("초기화: 카카오뱅크 터치시작");
                    TouchInjector.TouchDrag(new (int x, int y)[]
                            {
                            (1622, 543),
                            (1822, 543),
                            (1617, 722),
                            (1822, 722)
                            }, 7, 20);


                    if (token.WaitHandle.WaitOne(5000)) return;

                }
                else if (TouchInjector.IsColorMatch(1673, 800, Color.FromArgb(255, 255, 255))) // 팝업떴음
                {
                    if (token.WaitHandle.WaitOne(2000)) return;
                    Logger.Instance.Log("초기화: 팝업 닫기" + TouchInjector.getColor(1600, 803).ToString());
                    TouchInjector.TouchClick(1600, 803);
                    if (token.WaitHandle.WaitOne(2000)) return;
                }
                else if (TouchInjector.IsColorMatch(1709, 396, Color.FromArgb(254, 227, 0)) && TouchInjector.IsColorMatch(1709, 541, Color.FromArgb(254, 227, 0)) 
                    && TouchInjector.IsColorMatch(1709, 567, Color.FromArgb(236, 208, 194)) && TouchInjector.IsColorMatch(1733, 668, Color.FromArgb(236, 208, 194))) // 광고떴음
                {
                    if (token.WaitHandle.WaitOne(2000)) return;
                    Logger.Instance.Log("초기화: 팝업 닫기" + TouchInjector.getColor(1600, 803).ToString());
                    TouchInjector.TouchClick(1869, 176);
                    if (token.WaitHandle.WaitOne(2000)) return;
                }
                else if (TouchInjector.IsColorMatch(1823, 629, Color.FromArgb(236, 208, 194)))
                {
                    if (token.WaitHandle.WaitOne(2000)) return;
                    Logger.Instance.Log("초기화: 카카오뱅크 환율화면 진입" + TouchInjector.getColor(1718, 486).ToString());
                    TouchInjector.TouchClick(1823, 629);
                    break;
                }
                else if (TouchInjector.IsColorMatch(1718, 486, Color.FromArgb(236, 208, 194)))
                {
                    if (token.WaitHandle.WaitOne(2000)) return;
                    Logger.Instance.Log("초기화: 카카오뱅크 환율화면 진입" + TouchInjector.getColor(1718, 486).ToString());
                    TouchInjector.TouchClick(1718, 486);
                    break;
                }
                else if (TouchInjector.IsColorMatch(1718, 414, Color.FromArgb(236, 208, 194)))
                {
                    if (token.WaitHandle.WaitOne(2000)) return;
                    Logger.Instance.Log("초기화: 카카오뱅크 환율화면 진입" + TouchInjector.getColor(1718, 486).ToString());
                    TouchInjector.TouchClick(1718, 414);
                    break;
                }
                if (token.WaitHandle.WaitOne(500)) return;
            }

            while (true)
            {
                if (token.IsCancellationRequested) return;

                Logger.Instance.Log("초기화: 카카오뱅크 환율 기다리는중" + TouchInjector.getColor(1887, 109).ToString());
                if (!TouchInjector.IsColorMatch(1887, 109, Color.FromArgb(255, 255, 255)))
                {
                    Logger.Instance.Log("초기화: 카카오뱅크 환율화면 떴다" + TouchInjector.getColor(1887, 109).ToString());
                    break;
                }
                if (token.WaitHandle.WaitOne(100)) return;
            }

            if (token.WaitHandle.WaitOne(1000)) return;

            Logger.Instance.Log("초기화: 카카오뱅크 끝");
            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (!TouchInjector.IsColorMatch(1573, 766, Color.FromArgb(254, 227, 1)))
                {
                    Logger.Instance.Log("초기화: 카카오뱅크 홈버튼" + TouchInjector.getColor(1583, 340).ToString());
                    TouchInjector.TouchClick(1721, 856);
                    if (token.WaitHandle.WaitOne(100)) return;
                }
                else break;
                if (token.WaitHandle.WaitOne(300)) return;
            }
        }

        private void initTossApp(CancellationToken token)
        {
            if (token.WaitHandle.WaitOne(1000)) return;
            Logger.Instance.Log("초기화: 토스 시작");

            if (!TouchInjector.IsColorMatch(1863, 783, Color.FromArgb(17, 170, 255)))
            {
                Logger.Instance.Log("초기화: 토스 검은색 화면인듯" + TouchInjector.getColor(1863, 783).ToString());
                while (true)
                {
                    if (token.IsCancellationRequested) return;

                    if (!TouchInjector.IsColorMatch(1863, 783, Color.FromArgb(17, 170, 255)))
                    {
                        Logger.Instance.Log("초기화: 토스 검은색 화면 빽버튼" + TouchInjector.getColor(1863, 783).ToString());
                        TouchInjector.TouchClick(1833, 856);
                    }
                    else
                    {
                        break;
                    }
                    if (token.WaitHandle.WaitOne(1000)) return;
                }
            }

            Logger.Instance.Log("초기화: 토스 시작버튼 클릭" + TouchInjector.getColor(1863, 783).ToString());

            TouchInjector.TouchClickWithColor(1863, 783, Color.FromArgb(17, 170, 255), token);


            if (TouchInjector.IsColorMatch(1585, 778, Color.FromArgb(254, 227, 0)))
            {
                while (true)
                {
                    if (token.IsCancellationRequested) return;

                    if (TouchInjector.IsColorMatch(1585, 778, Color.FromArgb(254, 227, 0)))
                    {
                        Logger.Instance.Log("초기화: 토스 시작버튼 반복" + TouchInjector.getColor(1585, 778).ToString());
                        TouchInjector.TouchClick(1863, 783);
                    }
                    else
                    {
                        break;
                    }
                    if (token.WaitHandle.WaitOne(1000)) return;
                }
            }


            Logger.Instance.Log("초기화: 토스 전체 메뉴 클릭" + TouchInjector.getColor(1656, 190).ToString());
            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (TouchInjector.IsColorMatch(1656, 190, Color.FromArgb(242, 244, 245)))
                {
                    Logger.Instance.Log("초기화: 토스 전체 메뉴 클릭 반복 클릭" + TouchInjector.getColor(1656, 190).ToString());
                    TouchInjector.TouchClick(1860, 794);
                    break;
                }
                if (token.WaitHandle.WaitOne(1000)) return;
            }

            if (token.WaitHandle.WaitOne(3000)) return;


            Logger.Instance.Log("토스: 토스뱅크 환전하기 누름" + TouchInjector.getColor(1786, 183).ToString());
            TouchInjector.TouchClick(1554, 296);
            if (token.WaitHandle.WaitOne(100)) return;

            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (!TouchInjector.IsColorMatch(1887, 109, Color.FromArgb(255, 255, 255)))
                {
                    Logger.Instance.Log("토스: 토스뱅크 진입" + TouchInjector.getColor(1887, 109).ToString());
                    break;
                }
                if (token.WaitHandle.WaitOne(300)) return;
            }

            Logger.Instance.Log("초기화: 토스뱅크 환율 화면 진입" + TouchInjector.getColor(1887, 109).ToString());

            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (TouchInjector.IsColorMatch(1887, 109, Color.FromArgb(10, 15, 20)))
                {
                    Logger.Instance.Log("초기화: 토스뱅크 환율진입중" + TouchInjector.getColor(1887, 109).ToString());
                    break;
                }

                Logger.Instance.Log("초기화: 토스뱅크 환율 화면 기다리는중" + TouchInjector.getColor(1887, 109).ToString());


                if (token.WaitHandle.WaitOne(400)) return;
            }

            if (token.WaitHandle.WaitOne(1700)) return;

            Logger.Instance.Log("초기화: 토스 끝");
            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (!TouchInjector.IsColorMatch(1573, 766, Color.FromArgb(254, 227, 1)))
                {
                    Logger.Instance.Log("초기화: 토스 홈버튼" + TouchInjector.getColor(1583, 340).ToString());
                    TouchInjector.TouchClick(1721, 856);
                    if (token.WaitHandle.WaitOne(100)) return;
                }
                else break;
                if (token.WaitHandle.WaitOne(300)) return;
            }
        }

        private void initSwitchwonApp(CancellationToken token)
        {
            Logger.Instance.Log("초기화: 스위치원 시작" + TouchInjector.getColor(1790, 775).ToString());
            TouchInjector.TouchClickWithColor(1790, 775, Color.FromArgb(249, 169, 72), token);
            if (TouchInjector.IsColorMatch(1585, 778, Color.FromArgb(254, 227, 0)))
            {
                Logger.Instance.Log("초기화: 스위치원 화면 미표시" + TouchInjector.getColor(1790, 775).ToString());
                while (true)
                {
                    if (token.IsCancellationRequested) return;

                    if (TouchInjector.IsColorMatch(1585, 778, Color.FromArgb(254, 227, 0)))
                    {
                        Logger.Instance.Log("초기화: 스위치원 시작 다시 클릭" + TouchInjector.getColor(1790, 775).ToString());
                        TouchInjector.TouchClick(1790, 775);
                    }
                    else
                    {
                        break;
                    }
                    if (token.WaitHandle.WaitOne(3000)) return;
                }
            }

            if (token.WaitHandle.WaitOne(1000)) return;

            Logger.Instance.Log("초기화: 스위치원  환율 화면" + TouchInjector.getColor(1583, 340).ToString());
            while (true)
            {
                if (token.IsCancellationRequested) return;

                // Thread.Sleep(5000);
                // Logger.Instance.Log(TouchInjector.getColor(1583, 340).ToString());
                if (TouchInjector.IsColorMatch(1803, 495, Color.FromArgb(25, 35, 51)))
                {
                    Logger.Instance.Log("초기화: 스위치원  광고 닫기1");
                    TouchInjector.TouchClick(1803, 495);
                }
                else if (TouchInjector.IsColorMatch(1803, 529, Color.FromArgb(25, 35, 51)))
                {
                    Logger.Instance.Log("초기화: 스위치원 광고 닫기2");
                    TouchInjector.TouchClick(1803, 529);
                }
                else if (TouchInjector.IsColorMatch(1870, 765, Color.FromArgb(25, 35, 51)))
                {
                    Logger.Instance.Log("초기화: 스위치원  광고 닫기3");
                    TouchInjector.TouchClick(1870, 765);
                }
                else if (TouchInjector.IsColorMatch(1583, 340, Color.FromArgb(67, 71, 77)))
                {
                    break;
                }
                if (token.WaitHandle.WaitOne(300)) return;
            }

            Logger.Instance.Log("초기화: 스위치원  끝");
            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (!TouchInjector.IsColorMatch(1573, 766, Color.FromArgb(254, 227, 1)))
                {
                    Logger.Instance.Log("초기화: 스위치원  홈버튼" + TouchInjector.getColor(1583, 340).ToString());
                    TouchInjector.TouchClick(1721, 856);
                    if (token.WaitHandle.WaitOne(100)) return;
                }
                else break;
                if (token.WaitHandle.WaitOne(300)) return;
            }
        }

        private async Task RunAutomationWebLoopAsync(CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                if (token.IsCancellationRequested) break;

                await updateHanaRate();

                // [중요] Thread.Sleep 대신 WaitHandle 사용
                // 5초 대기 중 취소되면 즉시 루프 종료
                if (token.WaitHandle.WaitOne(5000)) break;
            }

            Logger.Instance.Log("RunAutomationWebLoopAsync 루프 종료됨");
        }

        private async Task RunAutomationCheckErrorAsync()
        {
            // 이 함수는 ctsMain 토큰과 독립적으로 실행되는 감시자 스레드입니다.
            // 필요하다면 여기도 종료 로직을 개선할 수 있지만, 
            // 요청하신 "RunAutomationLoopAsync 하위"에 포함되지 않아 isRunning 플래그 유지합니다.
            while (isRunning)
            {

                try
                {
                    Bitmap bmp = new Bitmap(200, 150);
                    using (Graphics g = Graphics.FromImage(bmp))
                    {
                        g.CopyFromScreen(1540, 162, 0, 0, new Size(200, 150));
                    }

                    Thread.Sleep(2000);

                    Bitmap bmp2 = new Bitmap(200, 150);
                    using (Graphics g = Graphics.FromImage(bmp2))
                    {
                        g.CopyFromScreen(1540, 162, 0, 0, new Size(200, 150));
                    }

                    Thread.Sleep(3000);

                    Bitmap bmp3 = new Bitmap(200, 150);
                    using (Graphics g = Graphics.FromImage(bmp3))
                    {
                        g.CopyFromScreen(1540, 162, 0, 0, new Size(200, 150));
                    }

                    Thread.Sleep(2000);

                    Bitmap bmp4 = new Bitmap(200, 150);
                    using (Graphics g = Graphics.FromImage(bmp4))
                    {
                        g.CopyFromScreen(1540, 162, 0, 0, new Size(200, 150));
                    }

                    Thread.Sleep(3000);

                    Bitmap bmp5 = new Bitmap(200, 150);
                    using (Graphics g = Graphics.FromImage(bmp5))
                    {
                        g.CopyFromScreen(1540, 162, 0, 0, new Size(200, 150));
                    }

                    Thread.Sleep(1000);

                    Bitmap bmp6 = new Bitmap(200, 150);
                    using (Graphics g = Graphics.FromImage(bmp6))
                    {
                        g.CopyFromScreen(1540, 162, 0, 0, new Size(200, 150));
                    }

                    Thread.Sleep(2000);

                    Bitmap bmp7 = new Bitmap(200, 150);
                    using (Graphics g = Graphics.FromImage(bmp7))
                    {
                        g.CopyFromScreen(1540, 162, 0, 0, new Size(200, 150));
                    }

                    Thread.Sleep(3000);

                    Bitmap bmp8 = new Bitmap(200, 150);
                    using (Graphics g = Graphics.FromImage(bmp8))
                    {
                        g.CopyFromScreen(1540, 162, 0, 0, new Size(200, 150));
                    }


                    double a1 = ImageSimilarity.CompareSimilarity((Bitmap)bmp, (Bitmap)bmp2);
                    double a2 = ImageSimilarity.CompareSimilarity((Bitmap)bmp, (Bitmap)bmp3);
                    double a3 = ImageSimilarity.CompareSimilarity((Bitmap)bmp, (Bitmap)bmp4);
                    double a4 = ImageSimilarity.CompareSimilarity((Bitmap)bmp, (Bitmap)bmp5);
                    double a5 = ImageSimilarity.CompareSimilarity((Bitmap)bmp, (Bitmap)bmp6);
                    double a6 = ImageSimilarity.CompareSimilarity((Bitmap)bmp, (Bitmap)bmp7);
                    double a7 = ImageSimilarity.CompareSimilarity((Bitmap)bmp, (Bitmap)bmp8);

                    bool shouldRestart = false;

                    this.Invoke((Delegate)(() =>
                    {
                        if (a1 <= 0.99 || a2 <= 0.99 || a3 <= 0.99 || a4 <= 0.99 || a5 <= 0.99 || a6 <= 0.99 || a7 <= 0.99)
                        {
                            errorCount = 0;
                        }
                        else
                        {
                            sendKakaotalkAnyMessage("최규수", "확인필요!!! ");
                            errorCount++;

                            if (errorCount >= 5)
                            {
                                Logger.Instance.Log("⚠ 연속 5회 검사 오류 → 메인 스레드 재시작");

                                needInitApp = true;
                                shouldRestart = true;
                                errorCount = 0;
                            }
                        }
                    }));

                    if (shouldRestart)
                    {
                        RestartMainThreads();
                    }

                    bmp.Dispose();
                    bmp2.Dispose();
                    bmp3.Dispose();
                    bmp4.Dispose();
                    bmp5.Dispose();
                    bmp6.Dispose();
                    bmp7.Dispose();
                    bmp8.Dispose();

                    Thread.Sleep(40000);
                }
                catch (Exception ex)
                {
                    Logger.Instance.Log("감시 스레드 에러: " + ex.Message);
                }
            }
        }

        private void RestartMainThreads()
        {
            Logger.Instance.Log("메인 루프 재시작 시작");

            // 1) 기존 루프 취소
            if (ctsMain != null)
            {
                ctsMain.Cancel(); // 토큰 취소 -> 스레드 내부 WaitOne 해제 및 break 트리거

                // 중요: 스레드가 실제로 종료될 때까지 기다림 (WaitOne덕분에 금방 종료됨)
                if (threadMain != null && threadMain.IsAlive)
                {
                    threadMain.Join();
                }
                if (threadWeb != null && threadWeb.IsAlive)
                {
                    threadWeb.Join();
                }
                ctsMain.Dispose();
            }

            // 2) 잠시 대기
            Thread.Sleep(500);

            // 3) 새로운 토큰 생성
            ctsMain = new CancellationTokenSource();

            // 4) 메인 루프 재시작
            threadMain = new Thread(() => RunAutomationLoopAsync(ctsMain.Token).Wait())
            {
                IsBackground = true
            };
            threadMain.SetApartmentState(ApartmentState.STA);
            threadMain.Start();

            threadWeb = new Thread(() => RunAutomationWebLoopAsync(ctsMain.Token).Wait())
            {
                IsBackground = true
            };
            threadWeb.SetApartmentState(ApartmentState.STA);
            threadWeb.Start();

            Logger.Instance.Log("메인 루프 재시작 완료 (t3는 유지)");
        }

        private void initHome(CancellationToken token)
        {
            Stopwatch sw = Stopwatch.StartNew();   // 시작
            Logger.Instance.Log("시작: 홈으로 이동");
            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (!TouchInjector.IsColorMatch(1573, 766, Color.FromArgb(254, 227, 1)))
                {
                    Logger.Instance.Log("시작: 홈버튼" + TouchInjector.getColor(1573, 766).ToString());
                    TouchInjector.TouchClick(1721, 856);
                    if (token.WaitHandle.WaitOne(100)) return;
                }
                else break;
                if (token.WaitHandle.WaitOne(300)) return;
            }

            sw.Stop();
            Logger.Instance.Log("시작: sw.milliseconds: " + sw.ElapsedMilliseconds);
        }

        private async Task updateHanaRate()
        {
            this.Invoke((Delegate)(() =>
            {
                webView21.Reload();
            }));

            await Task.Delay(1000);

            string investDValue = await InvestWaitAndGetCellAsync("1", "9"); // 달러


            string dValue = await HanaWaitAndGetCellAsync("1", "9"); // 달러
            string yValue = await HanaWaitAndGetCellAsync("2", "9"); // 엔화
            string tDValue = await HanaWaitAndGetCellAsync("7", "9"); // 대만달러
            string tbValue = await HanaWaitAndGetCellAsync("6", "9"); // 태국 바트
            string inValue = await HanaWaitAndGetCellAsync("37", "9"); // 인도네시아
            string vdValue = await HanaWaitAndGetCellAsync("11", "9"); // 베트남

            this.Invoke((Delegate)(() =>
            {
                hanaDLabel.Text = dValue.Replace(",", "");
                hanaYLabel.Text = yValue.Replace(",", "");
                hanaTDLabel.Text = tDValue.Replace(",", "");
                hanaTBLabel.Text = tbValue.Replace(",", "");
                hanaINLabel.Text = inValue.Replace(",", "");
                hanaVDLabel.Text = vdValue.Replace(",", "");
                investDLabel.Text = investDValue.Replace(",", "");
            }));
        }

        private void getRateFromKakaoBank(CancellationToken token)
        {
            Stopwatch sw = Stopwatch.StartNew();   // 시작
            Logger.Instance.Log("카카오뱅크: 시작" + TouchInjector.getColor(1585, 778).ToString());
            TouchInjector.TouchClickWithColor(1585, 778, Color.FromArgb(254, 227, 0), token);

            if (TouchInjector.IsColorMatch(1863, 783, Color.FromArgb(14, 168, 255)))
            {
                Logger.Instance.Log("카카오뱅크: 시작안됨" + TouchInjector.getColor(1863, 783).ToString());
                while (true)
                {
                    if (token.IsCancellationRequested) return;

                    if (TouchInjector.IsColorMatch(1863, 783, Color.FromArgb(17, 170, 255)))
                    {
                        Logger.Instance.Log("카카오뱅크: 시작 다시 클릭" + TouchInjector.getColor(1863, 783).ToString());
                        TouchInjector.TouchClick(1585, 778);
                    }
                    else break;
                    if (token.WaitHandle.WaitOne(300)) return;
                }
            }


            Logger.Instance.Log("카카오뱅크: 화면 진입" + TouchInjector.getColor(1850, 558).ToString());
            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (TouchInjector.IsColorMatch(1850, 558, Color.FromArgb(99, 110, 215)))
                {
                    if (token.WaitHandle.WaitOne(500)) return;
                    Logger.Instance.Log("카카오뱅크: 현재환율 클릭" + TouchInjector.getColor(1585, 778).ToString());
                    TouchInjector.TouchClick(1722, 624);
                    if (token.WaitHandle.WaitOne(500)) return;

                    break;
                }
                if (token.WaitHandle.WaitOne(300)) return;
            }

            Logger.Instance.Log("카카오뱅크: 현재환율 화면 진입" + TouchInjector.getColor(1766, 720).ToString());
            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (TouchInjector.IsColorMatch(1790, 540, Color.FromArgb(254, 227, 0)))
                {
                    Logger.Instance.Log("카카오뱅크: 에러화면 발생" + TouchInjector.getColor(1790, 540).ToString());
                    TouchInjector.TouchClick(1790, 540);

                    if (token.WaitHandle.WaitOne(1000)) return;

                    TouchInjector.TouchClick(1722, 624);
                }
                else if (TouchInjector.IsColorMatch(1766, 720, Color.FromArgb(244, 244, 244)))
                {
                    break;
                }
                if (token.WaitHandle.WaitOne(300)) return;
            }

            Logger.Instance.Log("카카오뱅크: 현재환율 화면");

            Bitmap bmp = new Bitmap(170, 50);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(1634, 198, 0, 0, new Size(170, 50));
            }

            string exchangeRate = OcrHelper.Instance.RecognizeEnglish(bmp);

            this.Invoke((Delegate)(() =>
            {
                pictureBox1.Image?.Dispose();
                pictureBox1.Image = (Bitmap)bmp.Clone();

                if (exchangeRate == "") return;

                double gap = Double.Parse(exchangeRate) - Double.Parse(CBCurRateLabel.Text);
                if (gap != 0)
                    CBCurGapLabel.Text = "(" + (gap > 0 ? "🔺" : "⬇️") + Math.Round(Math.Abs(gap), 2) + ")";
                else CBCurGapLabel.Text = "";


                CBCurRateLabel.Text = exchangeRate.Replace(",", "");

                double curRateValue = -1;
                double lowRateValue = -1;
                double highRateValue = -1;

                Double.TryParse(CBCurRateLabel.Text, out curRateValue);

                if (Double.TryParse(CBHighTextbox.Text, out highRateValue))
                {

                    if (curRateValue > highRateValue)
                    {
                        this.Invoke((Delegate)(() =>
                        {
                            CBHighTextbox.Text = CBCurRateLabel.Text;
                        }));
                    }
                }
                else
                {
                    this.Invoke((Delegate)(() =>
                    {
                        CBHighTextbox.Text = CBCurRateLabel.Text;
                    }));
                }
                if (Double.TryParse(CBLowTextbox.Text, out lowRateValue))
                {
                    if (curRateValue > 0 && curRateValue < lowRateValue)
                    {
                        this.Invoke((Delegate)(() =>
                        {
                            CBLowTextbox.Text = CBCurRateLabel.Text;
                        }));
                    }
                }
                else
                {
                    this.Invoke((Delegate)(() =>
                    {
                        CBLowTextbox.Text = CBCurRateLabel.Text;
                    }));
                }
            }));

            Logger.Instance.Log("카카오뱅크: 환율 인식 끝");


            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (TouchInjector.IsColorMatch(1754, 471, Color.FromArgb(236, 208, 194)))
                {
                    Logger.Instance.Log("카카오뱅크: 달러박스 밖으로 나와버렸음" + TouchInjector.getColor(1754, 471).ToString());
                    TouchInjector.TouchClick(1754, 471);
                    if (token.WaitHandle.WaitOne(900)) return;
                }
                else if (!TouchInjector.IsColorMatch(1851, 519, Color.FromArgb(99, 110, 215)))
                {
                    Logger.Instance.Log("카카오뱅크: <- 클릭" + TouchInjector.getColor(1851, 519).ToString());
                    TouchInjector.TouchClick(1525, 113);
                    if (token.WaitHandle.WaitOne(500)) return;
                }

                else break;

                if (token.WaitHandle.WaitOne(300)) return;
            }

            Logger.Instance.Log("카카오뱅크: 달러박스 화면" + TouchInjector.getColor(1851, 519).ToString());

            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (!TouchInjector.IsColorMatch(1585, 778, Color.FromArgb(254, 227, 0)))
                {
                    Logger.Instance.Log("카카오뱅크: 홈버튼" + TouchInjector.getColor(1585, 778).ToString());
                    TouchInjector.TouchClick(1721, 856);
                    if (token.WaitHandle.WaitOne(200)) return;
                }
                else break;

                if (token.WaitHandle.WaitOne(300)) return;

            }
            bmp.Dispose();

            sw.Stop();
            Logger.Instance.Log("카뱅: sw.milliseconds: " + sw.ElapsedMilliseconds);
        }

        private void getRateFromToss(CancellationToken token)
        {
            Stopwatch sw = Stopwatch.StartNew();   // 시작
            Logger.Instance.Log("토스: 시작");
            if (!TouchInjector.IsColorMatch(1863, 783, Color.FromArgb(17, 170, 255)))
            {
                Logger.Instance.Log("토스: 검은색 화면인듯" + TouchInjector.getColor(1863, 783).ToString());
                while (true)
                {
                    if (token.IsCancellationRequested) return;

                    if (!TouchInjector.IsColorMatch(1863, 783, Color.FromArgb(17, 170, 255)))
                    {
                        Logger.Instance.Log("토스: 검은색 화면 빽버튼" + TouchInjector.getColor(1863, 783).ToString());
                        TouchInjector.TouchClick(1833, 856);
                    }
                    else
                    {
                        break;
                    }
                    if (token.WaitHandle.WaitOne(300)) return;
                }
            }

            Logger.Instance.Log("토스: 시작버튼 클릭" + TouchInjector.getColor(1863, 783).ToString());

            TouchInjector.TouchClickWithColor(1863, 783, Color.FromArgb(17, 170, 255), token);

            if (token.WaitHandle.WaitOne(300)) return;

            if (TouchInjector.IsColorMatch(1886, 108, Color.FromArgb(255, 255, 255)))
            {
                while (true)
                {
                    if (token.IsCancellationRequested) return;

                    if (TouchInjector.IsColorMatch(1886, 108, Color.FromArgb(255, 255, 255)))
                    {
                        Logger.Instance.Log("토스: 시작버튼 반복" + TouchInjector.getColor(1886, 108).ToString());
                        TouchInjector.TouchClick(1863, 783);
                    }
                    else
                    {
                        break;
                    }
                    if (token.WaitHandle.WaitOne(300)) return;
                }
            }

            Logger.Instance.Log("토스: 뱅크 밖으로 나갔는지 체크중" + TouchInjector.getColor(1786, 183).ToString());
            if (TouchInjector.IsColorMatch(1786, 183, Color.FromArgb(242, 244, 245)))
            {
                Logger.Instance.Log("토스: 토스뱅크 환전하기 누름" + TouchInjector.getColor(1786, 183).ToString());
                TouchInjector.TouchClick(1554, 296);
                if (token.WaitHandle.WaitOne(100)) return;

                while (true)
                {
                    if (token.IsCancellationRequested) return;

                    if (!TouchInjector.IsColorMatch(1887, 109, Color.FromArgb(255, 255, 255)))
                    {
                        Logger.Instance.Log("토스: 토스뱅크 진입" + TouchInjector.getColor(1887, 109).ToString());
                        break;
                    }
                    if (token.WaitHandle.WaitOne(300)) return;
                }
            }


            if (token.WaitHandle.WaitOne(300)) return;
            Logger.Instance.Log("토스: 아래로 스크롤");
            TouchInjector.TouchDrag(new (int x, int y)[]
                        {
                                (1714, 774),
                                (1714, 354),
                        }, 5, 100);
            if (token.WaitHandle.WaitOne(1000)) return;

            updateExchageRate(@"image\대만.png", pictureBox7, tossTDLabel);
            updateExchageRate(@"image\태국.png", pictureBox8, tossTBLabel);
            updateExchageRate(@"image\인도네시아.png", pictureBox9, tossINLabel);

            Logger.Instance.Log("토스: 위로 스크롤");
            TouchInjector.TouchDrag(new (int x, int y)[]
                        {
                                (1714, 354),
                                (1714, 800),
                        }, 5, 100);

            TouchInjector.TouchDrag(new (int x, int y)[]
                        {
                                (1714, 354),
                                (1714, 800),
                        }, 5, 100);


            updateExchageRate(@"image\미국.png", pictureBox4, tossDLabel);
            updateExchageRate(@"image\일본.png", pictureBox5, tossYLabel);
            updateExchageRate(@"image\베트남.png", pictureBox6, tossVDLabel);

            Logger.Instance.Log("토스: 끝" + TouchInjector.getColor(1585, 778).ToString());

            //Logger.Instance.Log("토스: 오류발생" + TouchInjector.getColor(1900, 56).ToString()); // 208 208 208
            //TouchInjector.TouchClickWithColor(1563, 736, Color.FromArgb(4, 83, 109));

            //Logger.Instance.Log("토스: 전체 메뉴 클릭" + TouchInjector.getColor(1861, 790).ToString());
            //TouchInjector.TouchClickWithColor(1861, 790, Color.FromArgb(203, 208, 210));

            sw.Stop();
            Logger.Instance.Log("토스: sw.milliseconds: " + sw.ElapsedMilliseconds);
        }

        private void updateExchageRate(string path, PictureBox pictureBox, Label label)
        {
            var pos = ImageFinder.FindImageOnScreen(path, 0.85);

            Logger.Instance.Log("토스: 이미지 위치 찾음: " + path + " " + (pos != null ? pos.Value.ToString() : ""));

            if (pos != null)
            {
                Bitmap bmp = new Bitmap(110, 30);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(1787, pos.Value.Y - 20, 0, 0, new Size(110, 30));
                }

                for (int y = 0; y < bmp.Height; y++)
                {
                    for (int x = 0; x < bmp.Width; x++)
                    {
                        Color c = bmp.GetPixel(x, y);

                        if (c.R > 200 || c.B > 200)   // Red 값이 200 이상이면
                        {
                            bmp.SetPixel(x, y, Color.White); // 픽셀을 흰색으로 변경
                        }
                    }
                }

                string exchangeDRate = OcrHelper.Instance.RunOcr(bmp);
                Logger.Instance.Log("토스: 글자 인식 " + exchangeDRate);
                //}

                this.Invoke((Delegate)(() =>
                {
                    pictureBox.Image?.Dispose();
                    pictureBox.Image = (Bitmap)bmp.Clone();
                    label.Text = exchangeDRate.Replace(",", "");
                }));
            }
        }

        private void getRateFromKakaoPay(CancellationToken token)
        {
            Stopwatch sw = Stopwatch.StartNew();   // 시작
            Logger.Instance.Log("카페: 시작" + TouchInjector.getColor(1713, 773).ToString());
            TouchInjector.TouchClickWithColor(1713, 773, Color.FromArgb(255, 235, 0), token);
            if (TouchInjector.IsColorMatch(1585, 778, Color.FromArgb(254, 227, 0)))
            {
                Logger.Instance.Log("카페: 화면 미표시" + TouchInjector.getColor(1713, 773).ToString());
                while (true)
                {
                    if (token.IsCancellationRequested) return;

                    if (TouchInjector.IsColorMatch(1585, 778, Color.FromArgb(254, 227, 0)))
                    {
                        Logger.Instance.Log("스위치: 시작 다시 클릭" + TouchInjector.getColor(1713, 773).ToString());
                        TouchInjector.TouchClick(1713, 773);
                    }
                    else
                    {
                        break;
                    }
                    if (token.WaitHandle.WaitOne(3000)) return;
                }
            }


            Logger.Instance.Log("카페: 원->달러");
            if (token.WaitHandle.WaitOne(800)) return;
            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (TouchInjector.IsColorMatch(1806, 308, Color.FromArgb(254, 61, 76))) break;
                if (token.WaitHandle.WaitOne(100)) return;
            }

            Bitmap bmp = new Bitmap(167, 60);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(1537, 433, 0, 0, new Size(167, 60));
            }

            string exchangeDWRate = OcrHelper.Instance.RecognizeEnglish(bmp);


            this.Invoke((Delegate)(() =>
            {
                pictureBox2.Image?.Dispose();
                pictureBox2.Image = (Bitmap)bmp.Clone();

                if (exchangeDWRate == "") return;

                CPCurDWRateLabel.Text = exchangeDWRate.Replace(",", "");
            }));


            if (token.WaitHandle.WaitOne(800)) return;
            Logger.Instance.Log("카페: 달러->원");
            TouchInjector.TouchClick(1577, 155);


            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (TouchInjector.IsColorMatch(1829, 391, Color.FromArgb(255, 63, 87))) break;
                if (token.WaitHandle.WaitOne(100)) return;
            }

            Bitmap bmp1 = new Bitmap(167, 60);
            using (Graphics g = Graphics.FromImage(bmp1))
            {
                g.CopyFromScreen(1537, 433, 0, 0, new Size(167, 60));
            }

            string exchangeWDRate = OcrHelper.Instance.RecognizeEnglish(bmp1);


            this.Invoke((Delegate)(() =>
            {
                pictureBox3.Image?.Dispose();
                pictureBox3.Image = (Bitmap)bmp1.Clone();

                if (exchangeWDRate == "") return;

                CPCurWDRateLabel.Text = exchangeWDRate.Replace(",", "");
            }));

            if (token.WaitHandle.WaitOne(800)) return;
            TouchInjector.TouchClick(1824, 163);

            Logger.Instance.Log("카페: 끝");
            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (!TouchInjector.IsColorMatch(1573, 766, Color.FromArgb(254, 227, 1)))
                {
                    Logger.Instance.Log("카페: 홈버튼" + TouchInjector.getColor(1583, 340).ToString());
                    TouchInjector.TouchClick(1721, 856);
                    if (token.WaitHandle.WaitOne(100)) return;
                }
                else break;
                if (token.WaitHandle.WaitOne(300)) return;
            }


            sw.Stop();
            Logger.Instance.Log("카페: sw.milliseconds: " + sw.ElapsedMilliseconds);

            bmp.Dispose();
            bmp1.Dispose();
        }

        private void getRateFromSwitchwon(CancellationToken token)
        {
            Stopwatch sw = Stopwatch.StartNew();   // 시작
            Logger.Instance.Log("스위치: 시작" + TouchInjector.getColor(1790, 775).ToString());
            TouchInjector.TouchClickWithColor(1790, 775, Color.FromArgb(249, 169, 72), token);
            if (TouchInjector.IsColorMatch(1585, 778, Color.FromArgb(254, 227, 0)))
            {
                Logger.Instance.Log("스위치: 화면 미표시" + TouchInjector.getColor(1790, 775).ToString());
                while (true)
                {
                    if (token.IsCancellationRequested) return;

                    if (TouchInjector.IsColorMatch(1585, 778, Color.FromArgb(254, 227, 0)))
                    {
                        Logger.Instance.Log("스위치: 시작 다시 클릭" + TouchInjector.getColor(1790, 775).ToString());
                        TouchInjector.TouchClick(1790, 775);
                    }
                    else
                    {
                        break;
                    }
                    if (token.WaitHandle.WaitOne(300)) return;
                }
            }

            if (token.WaitHandle.WaitOne(1000)) return;

            Logger.Instance.Log("스위치: 환율 화면" + TouchInjector.getColor(1583, 340).ToString());
            while (true)
            {
                if (token.IsCancellationRequested) return;

                // Thread.Sleep(5000);
                // Logger.Instance.Log(TouchInjector.getColor(1583, 340).ToString());
                if (TouchInjector.IsColorMatch(1803, 495, Color.FromArgb(25, 35, 51)))
                {
                    Logger.Instance.Log("스위치: 광고 닫기1");
                    TouchInjector.TouchClick(1803, 495);
                }
                else if (TouchInjector.IsColorMatch(1803, 529, Color.FromArgb(25, 35, 51)))
                {
                    Logger.Instance.Log("스위치: 광고 닫기2");
                    TouchInjector.TouchClick(1803, 529);
                }
                else if (TouchInjector.IsColorMatch(1870, 765, Color.FromArgb(25, 35, 51)))
                {
                    Logger.Instance.Log("스위치: 광고 닫기3");
                    TouchInjector.TouchClick(1870, 765);
                }
                else if (TouchInjector.IsColorMatch(1583, 340, Color.FromArgb(67, 71, 77)))
                {
                    break;
                }
                if (token.WaitHandle.WaitOne(300)) return;
            }

            Logger.Instance.Log("스위치: 캡쳐");
            Bitmap bmp = new Bitmap(130, 35);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(1548, 287, 0, 0, new Size(130, 35));
            }

            for (int y = 0; y < bmp.Height; y++)
            {
                for (int x = 0; x < bmp.Width; x++)
                {
                    Color c = bmp.GetPixel(x, y);

                    if (c.R > 200)   // Red 값이 200 이상이면
                    {
                        bmp.SetPixel(x, y, Color.White); // 픽셀을 흰색으로 변경
                    }
                }
            }

            string exchangeDWRate = OcrHelper.Instance.RecognizeEnglish(bmp);


            this.Invoke((Delegate)(() =>
            {
                pictureBox10.Image?.Dispose();
                pictureBox10.Image = (Bitmap)bmp.Clone();

                if (exchangeDWRate == "") return;

                switchDLabel.Text = exchangeDWRate.Replace(",", "");
            }));

            Logger.Instance.Log("스위치: 끝");
            while (true)
            {
                if (token.IsCancellationRequested) return;

                if (!TouchInjector.IsColorMatch(1573, 766, Color.FromArgb(254, 227, 1)))
                {
                    Logger.Instance.Log("스위치: 홈버튼" + TouchInjector.getColor(1583, 340).ToString());
                    TouchInjector.TouchClick(1721, 856);
                    if (token.WaitHandle.WaitOne(100)) return;
                }
                else break;
                if (token.WaitHandle.WaitOne(300)) return;
            }


            bmp.Dispose();

            sw.Stop();
            Logger.Instance.Log("스위치: sw.milliseconds: " + sw.ElapsedMilliseconds);
        }

        private void sendKakaotalkAnyMessage(string roomName, string message)
        {
            IntPtr chattingRoom = FindWindow(null, roomName);
            IntPtr textBoxHwnd = FindWindowEx(chattingRoom, IntPtr.Zero, "RICHEDIT50W", null);



            sendToKakaotalk(roomName, message);
            Thread.Sleep(200);
            PostMessage(textBoxHwnd, WM_KEYDOWN, (IntPtr)VK_ENTER, IntPtr.Zero);
            PostMessage(textBoxHwnd, WM_KEYUP, (IntPtr)VK_ENTER, IntPtr.Zero);

        }

        private void sendKakaotalkMessage(string roomName)
        {
            Logger.Instance.Log("메시지 전송");
            IntPtr chattingRoom = FindWindow(null, roomName);
            IntPtr textBoxHwnd = FindWindowEx(chattingRoom, IntPtr.Zero, "RICHEDIT50W", null);

            string bankText = "";
            string payDWText = "";
            string payWDText = "";

            if (CBHighTextbox.Text == CBCurRateLabel.Text)
            {
                bankText = "(카고)";
            }
            if (CBLowTextbox.Text == CBCurRateLabel.Text)
            {
                bankText = bankText + "(카저)";
            }

            if (CPHighDWRateTextbox.Text == CPCurDWRateLabel.Text)
            {
                payDWText = "(패원고)";
            }
            if (CPLowDWRateTextbox.Text == CPCurDWRateLabel.Text)
            {
                payDWText = payDWText + "(패원저)";
            }

            if (CPHighWDRateTextbox.Text == CPCurWDRateLabel.Text)
            {
                payWDText = "(패달고)";
            }
            if (CPLowWDRateTextbox.Text == CPCurWDRateLabel.Text)
            {
                payWDText = payWDText + "(패달저)";
            }

            string now = DateTime.Now.ToString("HH시 mm분 ss초");

            double dGap = Double.Parse(hanaDLabel.Text) - Double.Parse(tossDLabel.Text);
            tossDGapLabel.Text = "(" + Math.Round(Math.Abs(dGap), 2) + ")";

            double yGap = Double.Parse(hanaYLabel.Text) - Double.Parse(tossYLabel.Text);
            tossYGapLabel.Text = "(" + Math.Round(Math.Abs(yGap), 2) + ")";

            double tdGap = Double.Parse(hanaTDLabel.Text) - Double.Parse(tossTDLabel.Text);
            tossTDGapLabel.Text = "(" + Math.Round(Math.Abs(tdGap), 2) + ")";

            double tbGap = Double.Parse(hanaTBLabel.Text) - Double.Parse(tossTBLabel.Text);
            tossTBGapLabel.Text = "(" + Math.Round(Math.Abs(tbGap), 2) + ")";

            double inGap = Double.Parse(hanaINLabel.Text) - Double.Parse(tossINLabel.Text);
            tossINGapLabel.Text = "(" + Math.Round(Math.Abs(inGap), 2) + ")";

            double vdGap = Double.Parse(hanaVDLabel.Text) - Double.Parse(tossVDLabel.Text);
            tossVDGapLabel.Text = "(" + Math.Round(Math.Abs(vdGap), 2) + ")";

            string multiLine = "";
            DateTime nowTime = DateTime.Now;
            if (false)
            // if (DateTime.Now.Hour >= 9 && DateTime.Now.Minute >= 5 || (DateTime.Now.Hour >= 10 && DateTime.Now.Hour < 16))
            {
                multiLine = $@"카달: {CBCurRateLabel.Text}{CBCurGapLabel.Text}{bankText} 
카저: {CBLowTextbox.Text}// 카고: {CBHighTextbox.Text}
카패달>원: {CPCurWDRateLabel.Text}/원>달: {CPCurDWRateLabel.Text}

인달:{investDLabel.Text}/스달:{switchDLabel.Text}
하달:{hanaDLabel.Text}/토달:{tossDLabel.Text}/{tossDGapLabel.Text} 
하엔:{hanaYLabel.Text}/토엔:{tossYLabel.Text}/{tossYGapLabel.Text}
하대:{hanaTDLabel.Text}/토대:{tossTDLabel.Text}/{tossTDGapLabel.Text}
하바:{hanaTBLabel.Text}/토바:{tossTBLabel.Text}/{tossTBGapLabel.Text}
하루:{hanaINLabel.Text}/토루:{tossINLabel.Text}/{tossINGapLabel.Text}
하동:{hanaVDLabel.Text}/토동:{tossVDLabel.Text}/{tossVDGapLabel.Text}";
            }
            else
            {
                multiLine = $@"카달: {CBCurRateLabel.Text}{CBCurGapLabel.Text}{bankText} 
카저: {CBLowTextbox.Text}// 카고: {CBHighTextbox.Text}

인달:{investDLabel.Text}/스달:{switchDLabel.Text}
하달:{hanaDLabel.Text}/토달:{tossDLabel.Text}/{tossDGapLabel.Text} 
하엔:{hanaYLabel.Text}/토엔:{tossYLabel.Text}/{tossYGapLabel.Text}
하대:{hanaTDLabel.Text}/토대:{tossTDLabel.Text}/{tossTDGapLabel.Text}
하바:{hanaTBLabel.Text}/토바:{tossTBLabel.Text}/{tossTBGapLabel.Text}
하루:{hanaINLabel.Text}/토루:{tossINLabel.Text}/{tossINGapLabel.Text}
하동:{hanaVDLabel.Text}/토동:{tossVDLabel.Text}/{tossVDGapLabel.Text}";
            }
            sendToKakaotalk(roomName, now + "\r\n" + multiLine);
            Thread.Sleep(200);
            PostMessage(textBoxHwnd, WM_KEYDOWN, (IntPtr)VK_ENTER, IntPtr.Zero);
            PostMessage(textBoxHwnd, WM_KEYUP, (IntPtr)VK_ENTER, IntPtr.Zero);

        }

        private void saveValue()
        {
            this.Invoke((Delegate)(() =>
            {
                Properties.Settings.Default.CBHighValue = CBHighTextbox.Text;
                Properties.Settings.Default.CBLowValue = CBLowTextbox.Text;
                Properties.Settings.Default.CPHighDWValue = CPHighDWRateTextbox.Text;
                Properties.Settings.Default.CPLowDWValue = CPLowDWRateTextbox.Text;
                Properties.Settings.Default.CPHighWDValue = CPHighWDRateTextbox.Text;
                Properties.Settings.Default.CPLowWDValue = CPLowWDRateTextbox.Text;
                Properties.Settings.Default.Save();
            }));
        }

        private void sendToKakaotalk(string windowName, string data)
        {
            IntPtr chattingRoom = FindWindow(null, windowName);
            IntPtr textBoxHwnd = FindWindowEx(chattingRoom, IntPtr.Zero, "RICHEDIT50W", null);

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            static extern IntPtr SendMessage(IntPtr hWnd, int Msg, IntPtr wParam, string lParam);

            SendMessage(textBoxHwnd, EM_REPLACESEL, (IntPtr)1, data);

        }


        private async void MainForm_Load(object sender, EventArgs e)
        {
            string procName = "kakao_bank_macro"; // exe 이름에서 .exe 제거

            Process current = Process.GetCurrentProcess();
            Process[] processes = Process.GetProcessesByName(procName);

            foreach (Process p in processes)
            {
                try
                {
                    if (p.Id != current.Id)   // 자기 자신 제외
                    {
                        p.Kill();
                        p.WaitForExit();
                    }
                }
                catch { /* 프로세스 강제종료 실패해도 무시 */ }
            }

            CBHighTextbox.Text = Properties.Settings.Default.CBHighValue;
            CBLowTextbox.Text = Properties.Settings.Default.CBLowValue;
            CPHighDWRateTextbox.Text = Properties.Settings.Default.CPHighDWValue;
            CPLowDWRateTextbox.Text = Properties.Settings.Default.CPLowDWValue;
            CPHighWDRateTextbox.Text = Properties.Settings.Default.CPHighWDValue;
            CPLowWDRateTextbox.Text = Properties.Settings.Default.CPLowWDValue;

            await webView21.EnsureCoreWebView2Async(null);
            webView21.Source = new Uri("https://www.kebhana.com/cms/rate/index.do?contentUrl=/cms/rate/wpfxd651_01i.do#//HanaBank");

            await webView22.EnsureCoreWebView2Async(null);
            webView22.Source = new Uri("https://kr.investing.com/currencies/exchange-rates-table");

            Logger.Instance.Initialize(
        @"C:\log\log1.txt",
        (newText) =>
        {
            // UI 스레드에서 실행되도록 보장
            this.BeginInvoke(new Action(() =>
            {
                logTextBox.Text = newText;
            }));
        });
        }


        private async Task<string> HanaWaitAndGetCellAsync(string row, string col)
        {
            string script =
                "document.querySelector(\"table.tblBasic tbody tr:nth-child(" + row + ") td:nth-child(" + col + ")\")?.innerText";

            for (int i = 0; i < 30; i++)   // 최대 20번 * 300ms = 6초 대기
            {
                string result = await HanaExecuteJsAsync(script);
                string value = System.Text.Json.JsonSerializer.Deserialize<string>(result);

                if (!string.IsNullOrWhiteSpace(value))
                    return value;

                await Task.Delay(300);
            }

            return "0";
        }

        private Task<string> HanaExecuteJsAsync(string script)
        {
            var tcs = new TaskCompletionSource<string>();

            this.Invoke(new Action(async () =>
            {
                try
                {
                    string result = await webView21.ExecuteScriptAsync(script);
                    tcs.SetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            }));

            return tcs.Task;
        }

        private async Task<string> InvestWaitAndGetCellAsync(string row, string col)
        {
            string script =
                "document.querySelector(\"table tbody tr:nth-child(" + row + ") td:nth-child(" + col + ")\")?.innerText";

            for (int i = 0; i < 30; i++)   // 최대 20번 * 300ms = 6초 대기
            {
                string result = await InvestExecuteJsAsync(script);
                string value = System.Text.Json.JsonSerializer.Deserialize<string>(result);

                if (!string.IsNullOrWhiteSpace(value))
                    return value;

                await Task.Delay(300);
            }

            return "0";
        }

        private Task<string> InvestExecuteJsAsync(string script)
        {
            var tcs = new TaskCompletionSource<string>();

            this.Invoke(new Action(async () =>
            {
                try
                {
                    string result = await webView22.ExecuteScriptAsync(script);
                    tcs.SetResult(result);
                }
                catch (Exception ex)
                {
                    tcs.SetException(ex);
                }
            }));

            return tcs.Task;
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            Logger.Instance.Log("프로그램 종료 요청");

            isRunning = false;
            ctsMain?.Cancel();
            Thread.Sleep(300);

            Logger.Instance.Close();
        }

        private void MainForm_FormCornerPreferenceChanged(object sender, EventArgs e)
        {

        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            string procName = "kakao_bank_macro"; // exe 이름에서 .exe 제거

            Process current = Process.GetCurrentProcess();
            Process[] processes = Process.GetProcessesByName(procName);

            foreach (Process p in processes)
            {
                try
                {
                    if (p.Id != current.Id)   // 자기 자신 제외
                    {
                        p.Kill();
                        p.WaitForExit();
                    }
                }
                catch { /* 프로세스 강제종료 실패해도 무시 */ }
            }
        }

        private void webView22_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (webView22.CoreWebView2 == null)
                return;

            string script = @"
        document.querySelectorAll('video').forEach(v => v.remove());
    ";

            webView22.ExecuteScriptAsync(script);
        }
    }
}