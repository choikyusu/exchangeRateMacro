using Microsoft.Web.WebView2.Core;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace kakao_bank_macro
{
    public partial class MainForm : Form
    {
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



        public MainForm()
        {

            InitializeComponent();

        }

        private void startButton_Click(object sender, EventArgs e)
        {

            timerLabel.Text =  DateTime.Now.ToString("HH시 mm분 ss초");
            

            width = Screen.PrimaryScreen.Bounds.Width;

            IntPtr hWnd = FindWindow(null, "Samsung Flow");
            //IntPtr hWnd = FindWindow(null, "Galaxy S10");
            bool result = SetWindowPos(
            hWnd,
            HWND_TOPMOST,   // 항상 위로
            0, 0, 0, 0,     // 위치/크기 유지
            SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);


            if (!isRunning)
            {
                isRunning = true;
                Thread t = new Thread(async () => await RunAutomationLoopAsync());
                t.SetApartmentState(ApartmentState.STA);
                t.Start();

                Thread t2 = new Thread(async () => await RunAutomationWebLoopAsync());
                t2.SetApartmentState(ApartmentState.STA);
                t2.Start();

                Thread t3 = new Thread(async () => await RunAutomationCheckErrorAsync());
                t3.SetApartmentState(ApartmentState.STA);
                t3.Start();
            }
        }

        private async Task RunAutomationLoopAsync()
        {
            IntPtr hWnd = FindWindow(null, "Samsung Flow");
            // IntPtr hWnd = FindWindow(null, "Galaxy S10");
            this.Invoke((Delegate)(() =>
            {
                SetWindowPos(hWnd, IntPtr.Zero, width - 408, 0, 408, 900, SWP_NOZORDER | SWP_SHOWWINDOW);
            }));

            while (isRunning)
            {
                try
                {
                    var sw = System.Diagnostics.Stopwatch.StartNew(); // 시간 측정 시작
                    DateTime now = DateTime.Now;
                    if (DateTime.Now.Hour >= 7 && Properties.Settings.Default.startMorning != DateTime.Now.ToString("yyyy-MM-dd"))
                    {
                        Properties.Settings.Default.startMorning = DateTime.Now.ToString("yyyy-MM-dd");
                        CBHighTextbox.Text = "";
                        CBLowTextbox.Text = "";
                    }

                    initHome();


                    //if (DateTime.Now.Hour >= 9 && DateTime.Now.Minute >= 5 || (DateTime.Now.Hour >= 10 && DateTime.Now.Hour < 16))
                    //    getRateFromKakaoPay();
                    getRateFromSwitchwon();
                    getRateFromKakaoBank();
                    getRateFromToss();
                    this.Invoke((Delegate)(() =>
                    {
                         // sendKakaotalkMessage("최규수");
                        if (now.DayOfWeek >= DayOfWeek.Monday &&
                           now.Hour >= 7 || now.DayOfWeek >= DayOfWeek.Tuesday)
                        {
                            sendKakaotalkMessage("NEW환도박방");
                        }
                    }));


                    saveValue();

                    sw.Stop();
                    int elapsedMs = (int)sw.ElapsedMilliseconds;   // 걸린 시간(ms)
                    Logger.Instance.Log("루프 1회 도는데 " + elapsedMs + "ms 소요");
                    int targetMs = 20000; // 25초
                    int remain = targetMs - elapsedMs;
                    if (remain > 0)
                        Thread.Sleep(remain);   // 남은 시간 만큼 Sleep

                }
                catch(Exception ex)
                {
                    var msg = ex.InnerException?.ToString() ?? ex.ToString();
                    Task.Run(() => MessageBox.Show(msg)); // UI 스레드 강요 없음
                }

            }
        }

        private async Task RunAutomationWebLoopAsync()
        {
            while (isRunning)
            {
                await updateHanaRate();

                Thread.Sleep(5000);
            }
        }

        private async Task RunAutomationCheckErrorAsync()
        {
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

                    this.Invoke((Delegate)(() =>
                    {
                        if (a1 <= 0.99 || a2 <= 0.99 || a3 <= 0.99 || a4 <= 0.99 || a5 <= 0.99 || a6 <= 0.99 || a7 <= 0.99)
                        {
                        }
                        else
                        {
                            sendKakaotalkAnyMessage("최규수", "확인필요!!! ");
                        }
                    }));

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
                    MessageBox.Show(ex.Message.ToString());

                }
            }
        }

        private void initHome()
        {
            Stopwatch sw = Stopwatch.StartNew();   // 시작
            Logger.Instance.Log("시작: 홈으로 이동");
            while (true)
            {
                if (!TouchInjector.IsColorMatch(1573, 766, Color.FromArgb(254, 227, 1)))
                {
                    Logger.Instance.Log("시작: 홈버튼" + TouchInjector.getColor(1573, 766).ToString());
                    TouchInjector.TouchClick(1721, 856);
                    System.Threading.Thread.Sleep(100);
                }
                else break;
                System.Threading.Thread.Sleep(300);
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

        private void getRateFromKakaoBank()
        {
            Stopwatch sw = Stopwatch.StartNew();   // 시작
            Logger.Instance.Log("카카오뱅크: 시작" + TouchInjector.getColor(1585, 778).ToString());
            TouchInjector.TouchClickWithColor(1585, 778, Color.FromArgb(254, 227, 0));

            if (TouchInjector.IsColorMatch(1863, 783, Color.FromArgb(14, 168, 255)))
            {
                Logger.Instance.Log("카카오뱅크: 시작안됨" + TouchInjector.getColor(1863, 783).ToString());
                while (true)
                {
                    if (TouchInjector.IsColorMatch(1863, 783, Color.FromArgb(17, 170, 255)))
                    {
                        Logger.Instance.Log("카카오뱅크: 시작 다시 클릭" + TouchInjector.getColor(1863, 783).ToString());
                        TouchInjector.TouchClick(1585, 778);
                    }
                    else break;
                    System.Threading.Thread.Sleep(300);
                }
            }


            Logger.Instance.Log("카카오뱅크: 화면 진입" + TouchInjector.getColor(1850, 558).ToString());
            while (true)
            {
                if (TouchInjector.IsColorMatch(1850, 558, Color.FromArgb(99, 110, 215)))
                {
                    Logger.Instance.Log("카카오뱅크: 현재환율 클릭" + TouchInjector.getColor(1585, 778).ToString());
                    System.Threading.Thread.Sleep(700);
                    TouchInjector.TouchClick(1722, 624);
                    System.Threading.Thread.Sleep(200);

                    break;
                }
                System.Threading.Thread.Sleep(300);
            }

            Logger.Instance.Log("카카오뱅크: 현재환율 화면 진입" + TouchInjector.getColor(1766, 720).ToString());
            while (true)
            {
                if (TouchInjector.IsColorMatch(1790, 540, Color.FromArgb(254, 227, 0)))
                {
                    Logger.Instance.Log("카카오뱅크: 에러화면 발생" + TouchInjector.getColor(1790, 540).ToString());
                    TouchInjector.TouchClick(1790, 540);

                    System.Threading.Thread.Sleep(1000);

                    TouchInjector.TouchClick(1722, 624);
                }
                else if (TouchInjector.IsColorMatch(1766, 720, Color.FromArgb(244, 244, 244)))
                {

                    break;
                }
                System.Threading.Thread.Sleep(300);
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
                if (!TouchInjector.IsColorMatch(1851, 519, Color.FromArgb(99, 110, 215)))
                {
                    Logger.Instance.Log("카카오뱅크: <- 클릭" + TouchInjector.getColor(1851, 519).ToString());
                    System.Threading.Thread.Sleep(700);
                    TouchInjector.TouchClick(1547, 113);
                    System.Threading.Thread.Sleep(200);
                }
                else break;

                System.Threading.Thread.Sleep(300);

            }

            Logger.Instance.Log("카카오뱅크: 달러박스 화면" + TouchInjector.getColor(1851, 519).ToString());

            while (true)
            {
                if (!TouchInjector.IsColorMatch(1585, 778, Color.FromArgb(254, 227, 0)))
                {
                    Logger.Instance.Log("카카오뱅크: 홈버튼" + TouchInjector.getColor(1585, 778).ToString());
                    TouchInjector.TouchClick(1721, 856);
                    System.Threading.Thread.Sleep(200);
                }
                else break;

                System.Threading.Thread.Sleep(300);

            }
            bmp.Dispose();

            sw.Stop();
            Logger.Instance.Log("카뱅: sw.milliseconds: " + sw.ElapsedMilliseconds);
        }

        private void getRateFromToss()
        {
            Stopwatch sw = Stopwatch.StartNew();   // 시작
            Logger.Instance.Log("토스: 시작");
            if (!TouchInjector.IsColorMatch(1863, 783, Color.FromArgb(17, 170, 255)))
            {
                Logger.Instance.Log("토스: 검은색 화면인듯" + TouchInjector.getColor(1863, 783).ToString());
                while (true)
                {
                    if (!TouchInjector.IsColorMatch(1863, 783, Color.FromArgb(17, 170, 255)))
                    {
                        Logger.Instance.Log("토스: 검은색 화면 빽버튼" + TouchInjector.getColor(1863, 783).ToString());
                        TouchInjector.TouchClick(1833, 856);
                    }
                    else
                    {
                        break;
                    }
                    System.Threading.Thread.Sleep(300);
                }
            }

            Logger.Instance.Log("토스: 시작버튼 클릭" + TouchInjector.getColor(1863, 783).ToString());

            TouchInjector.TouchClickWithColor(1863, 783, Color.FromArgb(17, 170, 255));

            System.Threading.Thread.Sleep(300);

            if (TouchInjector.IsColorMatch(1886, 108, Color.FromArgb(255, 255, 255)))
            {
                while (true)
                {
                    if (TouchInjector.IsColorMatch(1886, 108, Color.FromArgb(255, 255, 255)))
                    {
                        Logger.Instance.Log("토스: 시작버튼 반복" + TouchInjector.getColor(1886, 108).ToString());
                        TouchInjector.TouchClick(1863, 783);
                    }
                    else
                    {
                        break;
                    }
                    System.Threading.Thread.Sleep(300);
                }
            }

            Logger.Instance.Log("토스: 뱅크 밖으로 나갔는지 체크중" + TouchInjector.getColor(1786, 183).ToString());
            if (TouchInjector.IsColorMatch(1786, 183, Color.FromArgb(242, 244, 245)))
            {
                Logger.Instance.Log("토스: 토스뱅크 환전하기 누름" + TouchInjector.getColor(1786, 183).ToString());
                TouchInjector.TouchClick(1554, 296);
                System.Threading.Thread.Sleep(100);

                while (true)
                {
                    if (!TouchInjector.IsColorMatch(1887, 109, Color.FromArgb(255, 255, 255)))
                    {
                        Logger.Instance.Log("토스: 토스뱅크 진입" + TouchInjector.getColor(1887, 109).ToString());
                        break;
                    }
                    System.Threading.Thread.Sleep(300);
                }
            }


            Thread.Sleep(300);
            Logger.Instance.Log("토스: 아래로 스크롤");
            TouchInjector.TouchDrag(new (int x, int y)[]
                        {
                                (1714, 774),
                                (1714, 354),
                        }, 5, 100);
            Thread.Sleep(300);

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

        private void updateExchageRate (string path, PictureBox pictureBox, Label label )
        {
            var pos = ImageFinder.FindImageOnScreen(path, 0.85);

            if (pos != null)
            {
                Bitmap bmp = new Bitmap(110, 30);
                using (Graphics g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(1787, pos.Value.Y - 20, 0, 0, new Size(110, 30));
                }

                string exchangeDRate = OcrHelper.Instance.RecognizeEnglish(bmp);

                if (exchangeDRate == "0")
                    exchangeDRate = OcrHelper.Instance.RunOcr(bmp);

                this.Invoke((Delegate)(() =>
                {
                    pictureBox.Image?.Dispose();
                    pictureBox.Image = (Bitmap)bmp.Clone();
                    label.Text = exchangeDRate.Replace(",", "");
                }));
            }
        }

        private void getRateFromKakaoPay()
        {
            Stopwatch sw = Stopwatch.StartNew();   // 시작
            Logger.Instance.Log("카페: 시작" + TouchInjector.getColor(1713, 773).ToString());
            TouchInjector.TouchClickWithColor(1713, 773, Color.FromArgb(255, 235, 0));
            if (TouchInjector.IsColorMatch(1585, 778, Color.FromArgb(254, 227, 0)))
            {
                Logger.Instance.Log("카페: 화면 미표시" + TouchInjector.getColor(1713, 773).ToString());
                while (true)
                {
                    if (TouchInjector.IsColorMatch(1585, 778, Color.FromArgb(254, 227, 0)))
                    {
                        Logger.Instance.Log("스위치: 시작 다시 클릭" + TouchInjector.getColor(1713, 773).ToString());
                        TouchInjector.TouchClick(1713, 773);
                    }
                    else
                    {
                        break;
                    }
                    System.Threading.Thread.Sleep(3000);
                }
            }


            Logger.Instance.Log("카페: 원->달러");
            System.Threading.Thread.Sleep(800);
            while (true)
            {
                if (TouchInjector.IsColorMatch(1806, 308, Color.FromArgb(254, 61, 76))) break;
                System.Threading.Thread.Sleep(100);
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


            System.Threading.Thread.Sleep(800);
            Logger.Instance.Log("카페: 달러->원");
            TouchInjector.TouchClick(1577, 155);


            while (true)
            {
                if (TouchInjector.IsColorMatch(1829, 391, Color.FromArgb(255, 63, 87))) break;
                System.Threading.Thread.Sleep(100);
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

            System.Threading.Thread.Sleep(800);
            TouchInjector.TouchClick(1824, 163);

            Logger.Instance.Log("카페: 끝");
            while (true)
            {
                if (!TouchInjector.IsColorMatch(1573, 766, Color.FromArgb(254, 227, 1)))
                {
                    Logger.Instance.Log("카페: 홈버튼" + TouchInjector.getColor(1583, 340).ToString());
                    TouchInjector.TouchClick(1721, 856);
                    System.Threading.Thread.Sleep(100);
                }
                else break;
                System.Threading.Thread.Sleep(300);
            }


            sw.Stop();
            Logger.Instance.Log("카페: sw.milliseconds: " + sw.ElapsedMilliseconds);

            bmp.Dispose();
            bmp1.Dispose();
        }

        private void getRateFromSwitchwon()
        {
            Stopwatch sw = Stopwatch.StartNew();   // 시작
            Logger.Instance.Log("스위치: 시작" + TouchInjector.getColor(1790, 775).ToString());
            TouchInjector.TouchClickWithColor(1790, 775, Color.FromArgb(249, 169, 72));
            if (TouchInjector.IsColorMatch(1585, 778, Color.FromArgb(254, 227, 0)))
            {
                Logger.Instance.Log("스위치: 화면 미표시" + TouchInjector.getColor(1790, 775).ToString());
                while (true)
                {
                    if (TouchInjector.IsColorMatch(1585, 778, Color.FromArgb(254, 227, 0)))
                    {
                        Logger.Instance.Log("스위치: 시작 다시 클릭" + TouchInjector.getColor(1790, 775).ToString());
                        TouchInjector.TouchClick(1790, 775);
                    }
                    else 
                    {
                        break; 
                    }
                    System.Threading.Thread.Sleep(3000);
                }
            }

            System.Threading.Thread.Sleep(1000);

            Logger.Instance.Log("스위치: 환율 화면" + TouchInjector.getColor(1583, 340).ToString());
            while (true)
            {
                // Thread.Sleep(5000);
                // Logger.Instance.Log(TouchInjector.getColor(1583, 340).ToString());
                if(TouchInjector.IsColorMatch(1803, 495, Color.FromArgb(25, 35, 51)))
                {
                    TouchInjector.TouchClick(1803, 495);
                }
                else if (TouchInjector.IsColorMatch(1803, 529, Color.FromArgb(25, 35, 51)))
                {
                    TouchInjector.TouchClick(1803, 529);
                }
                else if (TouchInjector.IsColorMatch(1870, 765, Color.FromArgb(25, 35, 51)))
                {
                    TouchInjector.TouchClick(1870, 765);
                }
                else if (TouchInjector.IsColorMatch(1583, 340, Color.FromArgb(67, 71, 77))) 
                {
                    break; 
                }
                System.Threading.Thread.Sleep(300);
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
                if (!TouchInjector.IsColorMatch(1573, 766, Color.FromArgb(254, 227, 1)))
                {
                    Logger.Instance.Log("스위치: 홈버튼" + TouchInjector.getColor(1583, 340).ToString());
                    TouchInjector.TouchClick(1721, 856);
                    System.Threading.Thread.Sleep(100);
                }
                else break;
                System.Threading.Thread.Sleep(300);
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
            Logger.Instance.Close();

            isRunning = false;
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
