using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace kakao_bank_macro
{
    public class Logger
    {
        private static readonly Lazy<Logger> _instance =
            new Lazy<Logger>(() => new Logger());

        public static Logger Instance => _instance.Value;

        private StreamWriter logWriter;
        private readonly Queue<string> uiQueue = new Queue<string>();
        private readonly object fileLock = new object();

        private Action<string> uiLogCallback; // UI에 전달할 델리게이트

        private Logger() { }

        /// <summary>
        /// 반드시 프로그램 시작 시 1회 호출
        /// </summary>
        public void Initialize(string filePath, Action<string> uiLogCallback)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(filePath));

            logWriter = new StreamWriter(filePath, append: true, Encoding.UTF8)
            {
                AutoFlush = false
            };

            this.uiLogCallback = uiLogCallback;
        }

        /// <summary>
        /// 로그 기록 + UI 업데이트
        /// </summary>
        public void Log(string message)
        {
            int id = Thread.CurrentThread.ManagedThreadId;

            string timeMsg = $"{DateTime.Now:HH:mm:ss.fff} - {id} - {message}";

            // 1) 파일 로그
            Task.Run(() =>
            {
                lock (fileLock)
                {
                    logWriter.WriteLine(timeMsg);
                }
            });

            // 2) UI 업데이트
            if (uiLogCallback != null)
            {
                lock (uiQueue)
                {
                    uiQueue.Enqueue(message);
                    if (uiQueue.Count > 10)
                        uiQueue.Dequeue();
                }

                uiLogCallback(string.Join(Environment.NewLine, uiQueue.Reverse()));
            }
        }

        public void Close()
        {
            lock (fileLock)
            {
                logWriter?.Flush();
                logWriter?.Close();
            }
        }
    }
}