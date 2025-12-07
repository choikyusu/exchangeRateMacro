using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace kakao_bank_macro
{
    public static class ImageFinder
    {
        /// <summary>
        /// 큰 이미지(bigImage)에 작은 이미지(smallImage)가 포함된 위치를 찾는다.
        /// </summary>
        public static Point? FindImagePosition(Bitmap bigImage, Bitmap smallImage, double threshold = 0.85)
        {
            using (Mat big = BitmapToMat(bigImage))
            using (Mat small = BitmapToMat(smallImage))
            using (Mat result = new Mat())
            {
                CvInvoke.MatchTemplate(big, small, result, TemplateMatchingType.CcoeffNormed);

                double minVal = 0, maxVal = 0;
                Point minLoc = Point.Empty, maxLoc = Point.Empty;

                CvInvoke.MinMaxLoc(result, ref minVal, ref maxVal, ref minLoc, ref maxLoc);

                if (maxVal >= threshold)
                    return maxLoc;

                return null;
            }
        }

        public static Point? FindImageOnScreen(string relativePath, double threshold = 0.85)
        {
            string fullPath = Path.Combine(Application.StartupPath, relativePath);

            if (!File.Exists(fullPath))
                throw new FileNotFoundException("파일 없음:", fullPath);

            using (Bitmap templateBmp = new Bitmap(fullPath))
            using (Bitmap screenBmp = CaptureScreen())
            {
                return FindImagePosition(screenBmp, templateBmp, threshold);
            }
        }

        private static Bitmap CaptureScreen()
        {
            Rectangle bounds = Screen.PrimaryScreen.Bounds;
            Bitmap bmp = new Bitmap(bounds.Width, bounds.Height);

            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, bounds.Size);
            }

            return bmp; // 호출부에서 using 처리
        }

        /// <summary>
        /// Bitmap → Mat 변환 시 완전 복사 (clone dispose 문제 해결)
        /// </summary>
        private static Mat BitmapToMat(Bitmap bmp)
        {
            using (Bitmap clone = new Bitmap(bmp.Width, bmp.Height, System.Drawing.Imaging.PixelFormat.Format24bppRgb))
            {
                using (Graphics g = Graphics.FromImage(clone))
                    g.DrawImage(bmp, new Rectangle(0, 0, clone.Width, clone.Height));
                return clone.ToMat();
            }
        }
    }
}