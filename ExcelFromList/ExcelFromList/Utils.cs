using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

namespace ExcelFromList
{
    public static class Utils
    {
        public static bool IsFileLocked(string fullFileName)
        {
            var file = new FileInfo(fullFileName);
            FileStream stream = null;

            try
            {
                if (File.Exists(file.FullName))
                {
                    stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                }
                else
                {
                    return false;
                }
            }
            catch (IOException)
            {
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            return false;
        }

        public static void WaitForFileReady(string fullFileName)
        {
            try
            {
                if (File.Exists(fullFileName))
                {
                    while (IsFileLocked(fullFileName))
                        Thread.Sleep(100);
                }
            }
            catch (Exception)
            {
                throw;
            }
        }

        [DllImport("msvcrt.dll", CallingConvention = CallingConvention.Cdecl)]
        static extern int memcmp(byte[] b1, byte[] b2, long count);
        public static bool ByteArrayCompare(byte[] b1, byte[] b2)
        {
            // Validate buffers are the same length.
            // This also ensures that the count does not exceed the length of either buffer.  
            return b1.Length == b2.Length && memcmp(b1, b2, b1.Length) == 0;
        }

        public static Image ResizeImage(Image image, int newSize)
        {
            Bitmap destImage;
            try
            {
                double imageWidth = image.Width;
                double imageHeight = image.Height;
                double aspect = 1;

                if (imageWidth == imageHeight)
                {
                    imageWidth = newSize;
                    imageHeight = newSize;
                }
                else
                {
                    if (imageWidth > imageHeight)
                    {
                        aspect = imageWidth / imageHeight;
                        imageWidth = newSize;
                        imageHeight = newSize / aspect;
                    }

                    if (imageHeight > imageWidth)
                    {
                        aspect = imageWidth / imageHeight;
                        imageHeight = newSize;
                        imageWidth = imageHeight * aspect;
                    }
                }

                destImage = new Bitmap((int)imageWidth, (int)imageHeight);
                Rectangle destRect = new Rectangle(0, 0, (int)imageWidth, (int)imageHeight);

                using (var graphics = Graphics.FromImage(destImage))
                {
                    graphics.CompositingMode = CompositingMode.SourceCopy;
                    graphics.CompositingQuality = CompositingQuality.HighQuality;
                    graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    graphics.SmoothingMode = SmoothingMode.HighQuality;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

                    using (var wrapMode = new ImageAttributes())
                    {
                        wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                        graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }

            return destImage;
        }

    }

}
