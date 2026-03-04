using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace docment_tools_client.Helpers
{
    public class ImageHelper
    {
        /// <summary>
        /// 身份证图片预处理（灰度化+锐化+自适应局部二值化+降噪+缩放+倾斜校正）
        /// </summary>
        public static Bitmap PreprocessIdCardImage(string imagePath)
        {
            using (var img = new Bitmap(imagePath))
            {
                // 1. 转为灰度图
                var grayImg = ConvertToGrayscale(img);

                // 2. 图像锐化（提升文字边缘清晰度）
                var sharpImg = SharpenImage(grayImg);

                // 3. 自适应局部二值化（解决明暗不均，替代固定阈值）
                var binarizedImg = AdaptiveBinarization(sharpImg, 15, 10);

                // 4. 降噪（中值滤波，去除孤立噪点）
                var denoisedImg = MedianFilter(binarizedImg, 3);

                // 5. 倾斜校正（解决照片倾斜导致的识别误差）
                var correctedImg = CorrectImageSkew(denoisedImg);

                // 6. 缩放（提升小图识别率，固定2倍缩放，优化插值算法）
                var scaledImg = ScaleImage(correctedImg, 2);

                return scaledImg;
            }
        }

        #region 私有辅助方法
        /// <summary>
        /// 转为灰度图
        /// </summary>
        private static Bitmap ConvertToGrayscale(Bitmap original)
        {
            var grayImg = new Bitmap(original.Width, original.Height);
            using (var g = Graphics.FromImage(grayImg))
            {
                var grayMatrix = new ColorMatrix(
                    new float[][]
                    {
                        new float[] {0.299f, 0.299f, 0.299f, 0, 0},
                        new float[] {0.587f, 0.587f, 0.587f, 0, 0},
                        new float[] {0.114f, 0.114f, 0.114f, 0, 0},
                        new float[] {0, 0, 0, 1, 0},
                        new float[] {0, 0, 0, 0, 1}
                    });
                var attributes = new ImageAttributes();
                attributes.SetColorMatrix(grayMatrix);
                g.DrawImage(original, new Rectangle(0, 0, original.Width, original.Height),
                            0, 0, original.Width, original.Height, GraphicsUnit.Pixel, attributes);
            }
            return grayImg;
        }

        /// <summary>
        /// 图像锐化（提升文字边缘对比度）
        /// </summary>
        /// <summary>
        /// 图像锐化（提升文字边缘对比度）
        /// </summary>
        private static Bitmap SharpenImage(Bitmap original)
        {
            var sharpenMatrix = new float[3, 3]
            {
        { -1, -1, -1 },
        { -1,  9, -1 },
        { -1, -1, -1 }
            };

            var sharpImg = new Bitmap(original.Width, original.Height);

            // 锁定原图像的位数据
            var originalData = original.LockBits(new Rectangle(0, 0, original.Width, original.Height),
                ImageLockMode.ReadOnly, PixelFormat.Format8bppIndexed);

            // 锁定目标图像的位数据
            var sharpData = sharpImg.LockBits(new Rectangle(0, 0, sharpImg.Width, sharpImg.Height),
                ImageLockMode.WriteOnly, PixelFormat.Format8bppIndexed);

            int stride = originalData.Stride;
            IntPtr originalPtr = originalData.Scan0;
            int bytes = Math.Abs(stride) * original.Height;
            byte[] originalValues = new byte[bytes];

            // 从原图像复制数据
            System.Runtime.InteropServices.Marshal.Copy(originalPtr, originalValues, 0, bytes);

            // 创建输出数组
            byte[] outputValues = new byte[bytes];

            // 应用锐化卷积核
            for (int y = 1; y < original.Height - 1; y++)
            {
                for (int x = 1; x < original.Width - 1; x++)
                {
                    float sum = 0;
                    for (int ky = -1; ky <= 1; ky++)
                    {
                        for (int kx = -1; kx <= 1; kx++)
                        {
                            int idx = (y + ky) * stride + (x + kx);
                            sum += originalValues[idx] * sharpenMatrix[ky + 1, kx + 1];
                        }
                    }
                    sum = Math.Clamp(sum, 0, 255);
                    int outputIdx = y * stride + x;
                    outputValues[outputIdx] = (byte)sum;
                }
            }

            // 将处理后的数据写入目标图像
            System.Runtime.InteropServices.Marshal.Copy(outputValues, 0, sharpData.Scan0, outputValues.Length);

            // 解锁位数据
            original.UnlockBits(originalData);
            sharpImg.UnlockBits(sharpData);

            return sharpImg;
        }


        /// <summary>
        /// 自适应局部二值化（解决固定阈值明暗不均问题）
        /// </summary>
        /// <param name="grayImg">灰度图</param>
        /// <param name="blockSize">局部块大小</param>
        /// <param name="offset">阈值偏移量</param>
        private static Bitmap AdaptiveBinarization(Bitmap grayImg, int blockSize, int offset)
        {
            var binarizedImg = new Bitmap(grayImg.Width, grayImg.Height);
            var bitmapData = grayImg.LockBits(new Rectangle(0, 0, grayImg.Width, grayImg.Height),
                ImageLockMode.ReadOnly, PixelFormat.Format8bppIndexed);
            int stride = bitmapData.Stride;
            IntPtr ptr = bitmapData.Scan0;
            int bytes = Math.Abs(stride) * grayImg.Height;
            byte[] rgbValues = new byte[bytes];
            System.Runtime.InteropServices.Marshal.Copy(ptr, rgbValues, 0, bytes);

            for (int y = 0; y < grayImg.Height; y++)
            {
                for (int x = 0; x < grayImg.Width; x++)
                {
                    // 计算局部块内的像素均值
                    int sum = 0;
                    int count = 0;
                    int halfBlock = blockSize / 2;
                    for (int ky = -halfBlock; ky <= halfBlock; ky++)
                    {
                        for (int kx = -halfBlock; kx <= halfBlock; kx++)
                        {
                            int nx = x + kx;
                            int ny = y + ky;
                            if (nx >= 0 && nx < grayImg.Width && ny >= 0 && ny < grayImg.Height)
                            {
                                sum += rgbValues[ny * stride + nx];
                                count++;
                            }
                        }
                    }
                    int avg = sum / count;
                    int currentPixel = rgbValues[y * stride + x];
                    // 自适应阈值：当前像素 < (局部均值 - 偏移量) 则为黑，否则为白
                    Color newColor = currentPixel < (avg - offset) ? Color.Black : Color.White;
                    binarizedImg.SetPixel(x, y, newColor);
                }
            }

            grayImg.UnlockBits(bitmapData);
            return binarizedImg;
        }

        /// <summary>
        /// 中值滤波降噪（去除孤立噪点）
        /// </summary>
        private static Bitmap MedianFilter(Bitmap original, int kernelSize)
        {
            var denoisedImg = new Bitmap(original.Width, original.Height);
            int halfKernel = kernelSize / 2;

            for (int y = 0; y < original.Height; y++)
            {
                for (int x = 0; x < original.Width; x++)
                {
                    List<int> pixelValues = new List<int>();
                    // 遍历核内像素
                    for (int ky = -halfKernel; ky <= halfKernel; ky++)
                    {
                        for (int kx = -halfKernel; kx <= halfKernel; kx++)
                        {
                            int nx = x + kx;
                            int ny = y + ky;
                            if (nx >= 0 && nx < original.Width && ny >= 0 && ny < original.Height)
                            {
                                pixelValues.Add(original.GetPixel(nx, ny).R);
                            }
                        }
                    }
                    // 取中值作为当前像素值
                    pixelValues.Sort();
                    int median = pixelValues[pixelValues.Count / 2];
                    denoisedImg.SetPixel(x, y, median < 128 ? Color.Black : Color.White);
                }
            }
            return denoisedImg;
        }

        /// <summary>
        /// 倾斜校正（解决照片倾斜导致的识别误差）
        /// </summary>
        private static Bitmap CorrectImageSkew(Bitmap original)
        {
            // 边缘检测获取轮廓
            var edgeDetector = new Bitmap(original.Width, original.Height);
            using (var g = Graphics.FromImage(edgeDetector))
            using (var pen = new Pen(Color.Black, 1))
            {
                g.Clear(Color.White);
                for (int y = 1; y < original.Height - 1; y++)
                {
                    for (int x = 1; x < original.Width - 1; x++)
                    {
                        var p = original.GetPixel(x, y);
                        var pX1 = original.GetPixel(x - 1, y);
                        var pY1 = original.GetPixel(x, y - 1);
                        int diffX = Math.Abs(p.R - pX1.R);
                        int diffY = Math.Abs(p.R - pY1.R);
                        if (diffX > 30 || diffY > 30)
                        {
                            edgeDetector.SetPixel(x, y, Color.Black);
                        }
                    }
                }
            }

            // 计算倾斜角度（仅处理水平倾斜）
            float angle = 0;
            int maxCount = 0;
            for (float testAngle = -5; testAngle <= 5; testAngle += 0.5f)
            {
                using (var rotated = new Bitmap(edgeDetector.Width, edgeDetector.Height))
                using (var g = Graphics.FromImage(rotated))
                {
                    g.Clear(Color.White);
                    g.TranslateTransform(rotated.Width / 2, rotated.Height / 2);
                    g.RotateTransform(testAngle);
                    g.TranslateTransform(-rotated.Width / 2, -rotated.Height / 2);
                    g.DrawImage(edgeDetector, new Rectangle(0, 0, rotated.Width, rotated.Height));

                    // 统计垂直线数量
                    int count = 0;
                    for (int x = 0; x < rotated.Width; x++)
                    {
                        for (int y = 0; y < rotated.Height; y++)
                        {
                            if (rotated.GetPixel(x, y).R == 0)
                            {
                                count++;
                                break;
                            }
                        }
                    }
                    if (count > maxCount)
                    {
                        maxCount = count;
                        angle = testAngle;
                    }
                }
            }

            // 旋转校正
            var correctedImg = new Bitmap(original.Width, original.Height);
            using (var g = Graphics.FromImage(correctedImg))
            {
                g.Clear(Color.White);
                g.TranslateTransform(correctedImg.Width / 2, correctedImg.Height / 2);
                g.RotateTransform(-angle);
                g.TranslateTransform(-correctedImg.Width / 2, -correctedImg.Height / 2);
                g.DrawImage(original, new Rectangle(0, 0, correctedImg.Width, correctedImg.Height));
            }
            return correctedImg;
        }

        /// <summary>
        /// 缩放图片
        /// </summary>
        private static Bitmap ScaleImage(Bitmap original, float scaleFactor)
        {
            int newWidth = (int)(original.Width * scaleFactor);
            int newHeight = (int)(original.Height * scaleFactor);
            var scaledImg = new Bitmap(newWidth, newHeight);
            using (var g = Graphics.FromImage(scaledImg))
            {
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                g.DrawImage(original, new Rectangle(0, 0, newWidth, newHeight),
                            new Rectangle(0, 0, original.Width, original.Height), GraphicsUnit.Pixel);
            }
            return scaledImg;
        }
        #endregion
    }
}