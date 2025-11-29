using Emgu.CV;
using Emgu.CV.CvEnum;
using Emgu.CV.Structure;
using Emgu.CV.Util;

namespace kakao_bank_macro
{
    internal static class ImageSimilarity
    {
        /// <summary>
        /// 두 비트맵의 히스토그램 유사도를 비교합니다.
        /// </summary>
        /// <param name="bmp1">첫 번째 이미지</param>
        /// <param name="bmp2">두 번째 이미지</param>
        /// <returns>0.0 ~ 1.0 사이의 값 (1.0에 가까울수록 유사함)</returns>
        public static double CompareSimilarity(Bitmap bmp1, Bitmap bmp2)
        {

            // 1. Bitmap을 Emgu.CV의 Image 객체로 변환 (BGR 컬러 공간)
            // Emgu.CV.Bitmap 패키지가 필요합니다.
            using (Image<Bgr, byte> img1 = bmp1.ToImage<Bgr, byte>())
            using (Image<Bgr, byte> img2 = bmp2.ToImage<Bgr, byte>())
            {
                // 2. HSV 색상 공간으로 변환 
                // (RGB보다 조명 변화에 덜 민감하여 유사도 비교에 더 유리함)
                using (Image<Hsv, byte> hsv1 = img1.Convert<Hsv, byte>())
                using (Image<Hsv, byte> hsv2 = img2.Convert<Hsv, byte>())
                {
                    // 3. 히스토그램 계산 및 비교
                    return GetHistogramCorrelation(hsv1, hsv2);
                }
            }
        }

        private static double GetHistogramCorrelation(Image<Hsv, byte> img1, Image<Hsv, byte> img2)
        {
            // 히스토그램 설정
            // H(Hue)는 0~180, S(Saturation)는 0~256 범위를 가짐
            int hBins = 50;
            int sBins = 60;
            int[] histSize = { hBins, sBins };
            float[] ranges = { 0, 180, 0, 256 };
            int[] channels = { 0, 1 }; // Hue와 Saturation 채널만 사용 (Value는 밝기라 제외하여 조명 영향 최소화)

            using (Mat hist1 = new Mat())
            using (Mat hist2 = new Mat())
            using (VectorOfMat vImg1 = new VectorOfMat())
            using (VectorOfMat vImg2 = new VectorOfMat())
            {
                vImg1.Push(img1.Mat);
                vImg2.Push(img2.Mat);

                // 히스토그램 계산
                CvInvoke.CalcHist(vImg1, channels, null, hist1, histSize, ranges, false);
                CvInvoke.CalcHist(vImg2, channels, null, hist2, histSize, ranges, false);

                // 히스토그램 정규화 (크기가 다른 이미지도 비교 가능하게 함)
                CvInvoke.Normalize(hist1, hist1, 0, 1, NormType.MinMax);
                CvInvoke.Normalize(hist2, hist2, 0, 1, NormType.MinMax);

                // 히스토그램 비교 (Correl 방식: 1.0은 완전 일치, 낮을수록 불일치)
                // 다른 방식: ChiSqr, Intersect, Bhattacharyya 등
                double result = CvInvoke.CompareHist(hist1, hist2, HistogramCompMethod.Correl);

                return result;
            }
        }
    }
}
