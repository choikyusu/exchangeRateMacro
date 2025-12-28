using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading; // Thread 사용을 위해 추가

namespace kakao_bank_macro
{
    public class Logger
    {
        private static readonly Lazy<Logger> _instance =
            new Lazy<Logger>(() => new Logger());

        public static Logger Instance => _instance.Value;

        private StreamWriter logWriter;
        // 동시성 문제를 해결하기 위한 리스트 (Queue보다 인덱스 접근이 용이)
        private readonly List<string> uiList = new List<string>();

        // 락 객체 하나로 통합 관리 (파일/UI 따로 관리해도 되지만 단순화)
        private readonly object _lock = new object();

        private Action<string> uiLogCallback;

        private Logger() { }

        public void Initialize(string filePath, Action<string> uiLogCallback)
        {
            try
            {
                string dir = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(dir))
                    Directory.CreateDirectory(dir);

                // AutoFlush = true로 설정하여 크래시 발생 시에도 로그 저장 보장
                logWriter = new StreamWriter(filePath, append: true, Encoding.UTF8)
                {
                    AutoFlush = true
                };

                this.uiLogCallback = uiLogCallback;
            }
            catch (Exception ex)
            {
                // 초기화 실패 시 처리 (디버그용)
                System.Diagnostics.Debug.WriteLine($"Logger 초기화 실패: {ex.Message}");
            }
        }

        public void Log(string message)
        {
            // 아직 초기화 안됐으면 무시
            if (logWriter == null) return;

            int id = Thread.CurrentThread.ManagedThreadId;
            string timeMsg = $"{DateTime.Now:HH:mm:ss.fff} - [{id}] {message}";

            // UI에 보여질 텍스트를 미리 저장할 변수
            string uiTextToDisplay = null;

            // [중요] 모든 자원 접근을 lock 안에서 수행
            lock (_lock)
            {
                try
                {
                    // 1) 파일 쓰기 (동기적으로 수행하여 순서 보장)
                    logWriter.WriteLine(timeMsg);

                    // 2) UI 리스트 업데이트
                    if (uiLogCallback != null)
                    {
                        // 리스트에 추가
                        uiList.Add(message);

                        // 10개 넘어가면 앞에서부터 삭제 (Queue처럼 동작)
                        if (uiList.Count > 10)
                        {
                            uiList.RemoveAt(0); // 0번 인덱스(가장 오래된 것) 삭제
                        }

                        // [핵심 수정] lock 안에서 문자열을 완성해서 나감
                        // 컬렉션 변경 에러 방지
                        // Reverse()를 사용하여 최신 로그가 위로 오게 함 (필요 시 순서 조정)
                        uiTextToDisplay = string.Join(Environment.NewLine, Enumerable.Reverse(uiList));
                    }
                }
                catch
                {
                    // 로그 쓰다가 에러나면 무시 (프로그램 멈춤 방지)
                }
            }

            // 3) UI 콜백 호출 (Lock 밖에서 수행하여 데드락 방지)
            if (uiTextToDisplay != null && uiLogCallback != null)
            {
                uiLogCallback.Invoke(uiTextToDisplay);
            }
        }

        public void Close()
        {
            lock (_lock)
            {
                logWriter?.Close();
                logWriter = null;
            }
        }
    }
}