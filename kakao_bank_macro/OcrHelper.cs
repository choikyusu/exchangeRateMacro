using System.Text.RegularExpressions;
using Tesseract;

namespace kakao_bank_macro
{
    internal sealed class OcrHelper
    {
        private static readonly Lazy<OcrHelper> instance =
            new Lazy<OcrHelper>(() => new OcrHelper());

        public static OcrHelper Instance => instance.Value;

        private readonly TesseractEngine engine;

        // ⭐ 싱글턴 생성자
        private OcrHelper()
        {
            string tessDataPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "tessdata");
            string lang = "kor";

            if (!Directory.Exists(tessDataPath))
                throw new Exception($"tessdata 폴더 없음: {tessDataPath}");

            string trainedDataFile = Path.Combine(tessDataPath, $"{lang}.traineddata");
            if (!File.Exists(trainedDataFile))
                throw new Exception($"{lang}.traineddata 파일 없음: {trainedDataFile}");

            // ⚡ 엔진 1회만 생성
            engine = new TesseractEngine(tessDataPath, lang, EngineMode.Default);
        }

        // ⭐ OCR 함수
        public string RecognizeEnglish(Bitmap bmp)
        {
            try
            {
                // Bitmap → Pix 변환
                using var ms = new MemoryStream();
                bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                ms.Position = 0;

                using var pix = Pix.LoadFromMemory(ms.ToArray());
                using var page = engine.Process(pix);

                string text = page.GetText().Trim();

                string newText = Regex.Replace(text, @"[^0-9\.]", "");

                if (string.IsNullOrWhiteSpace(newText))
                    return "0";

                if (double.TryParse(newText, out double value))
                {
                    if (value > 10000 || value < 0)
                        return "0";
                }
                else
                {
                    return "0";
                }
                ms.Dispose();
                pix.Dispose();
                page.Dispose();

                return newText;
            }
            catch (Exception ex)
            {
                var msg = ex.InnerException?.ToString() ?? ex.ToString();
                Task.Run(() => MessageBox.Show(msg));
                return "0";
            }
        }
    }
}