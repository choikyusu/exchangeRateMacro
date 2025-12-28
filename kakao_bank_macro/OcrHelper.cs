using OpenCvSharp;
using OpenCvSharp.Extensions;
using Sdcb.PaddleOCR;
using Sdcb.PaddleOCR.Models.Local;
using System.Reflection.Metadata;
using System.Text.RegularExpressions;
using Tesseract;
using BitmapConverter = OpenCvSharp.Extensions.BitmapConverter;


namespace kakao_bank_macro
{
    internal sealed class OcrHelper
    {
        private static readonly Lazy<OcrHelper> instance =
            new Lazy<OcrHelper>(() => new OcrHelper());

        private static readonly Lazy<PaddleOcrAll> _paddleOcr
            = new Lazy<PaddleOcrAll>(() =>
            {
                return new PaddleOcrAll(
                    LocalFullModels.EnglishV3
                );
            });

        public static OcrHelper Instance => instance.Value;

        private readonly TesseractEngine engine;

        private static PaddleOcrAll Paddle => _paddleOcr.Value;

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
                    if (value > 1000000 || value < 0)
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

        public string RunOcr(Bitmap bitmap)
        {
            try
            {


                Mat mat = BitmapConverter.ToMat(bitmap);

                // 4채널(BGRA/ARGB) → 3채널(BGR) 변환
                if (mat.Channels() == 4)
                {
                    Cv2.CvtColor(mat, mat, ColorConversionCodes.BGRA2BGR);
                }


                var result = Paddle.Run(mat);

                string output = string.Join("\n", result.Regions.Select(r => r.Text));

                string text = output.Trim();

                string newText = Regex.Replace(text, @"[^0-9\.]", "");

                if (string.IsNullOrWhiteSpace(newText))
                    return "0";

                if (double.TryParse(newText, out double value))
                {
                    if (value > 1000000 || value < 0)
                        return "0";
                }
                else
                {
                    return "0";
                }

                return newText;

            }
            catch (Exception e)
            {

            }

            return "";
        }
    }
}