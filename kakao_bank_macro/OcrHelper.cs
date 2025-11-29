using System;
using System.Drawing;
using System.Text.RegularExpressions;
using Tesseract;

namespace kakao_bank_macro
{
    internal class OcrHelper
    {
        public static string RecognizeEnglish(Bitmap bmp)
        {
            try
            {
                // tessdata 폴더 경로 (실행 파일 기준)
                string tessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
                string lang = "kor";

                // 1️⃣ tessdata 폴더 존재 체크
                if (!Directory.Exists(tessDataPath))
                    throw new Exception($"tessdata 폴더 없음: {tessDataPath}");

                // 2️⃣ 언어 파일 체크
                string trainedDataFile = Path.Combine(tessDataPath, $"{lang}.traineddata");
                if (!File.Exists(trainedDataFile))
                    throw new Exception($"{lang}.traineddata 파일 없음: {trainedDataFile}");

                // 3️⃣ OCR 엔진 생성
                using var engine = new TesseractEngine(tessDataPath, lang, EngineMode.Default);

                // 4️⃣ Bitmap → Pix 변환
                using var ms = new MemoryStream();
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                ms.Position = 0;

                using var pix = Pix.LoadFromMemory(ms.ToArray());
                using var page = engine.Process(pix);

                string text = page.GetText().Trim();

                // 5️⃣ 숫자와 .만 추출
                string newText = Regex.Replace(text, @"[^0-9\.]", "");

                // 숫자 없으면 실패로 간주
                if (string.IsNullOrWhiteSpace(newText))
                    return "0";

                // 값 범위 체크
                if (double.TryParse(newText, out double value))
                {
                    if (value > 10000 || value < 0)
                        return "0";
                }
                else
                {
                    return "0";
                }

                return newText;
            }
            catch (Exception ex)
            {
                var msg = ex.InnerException?.ToString() ?? ex.ToString();
                Task.Run(() => MessageBox.Show(msg)); // UI 스레드 강요 없음
                return "0";
            }
        }
    }
}
