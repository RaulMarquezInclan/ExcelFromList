using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;

namespace ExcelFromList
{
    public static class Utils
    {

        public static string NoImage = "iVBORw0KGgoAAAANSUhEUgAAAGQAAABkCAIAAAD/gAIDAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAWBSURBVHhe7Zq/b9pQEMdNh0jZUilbOmSADDTqHvgLgKEeKtZuzghLlSXqULE0XcIYNlYrAx2Av6DuHjUZCkOGskUKW6Us9N7d8y9sE18CpkH3kdXGl4ff8zd3z/Dlcvd/eoaQjlf6fyEFIhYDEYuBiMVAxGIgYjEQsRiIWAxELAYiFgMRi4GIxUDEYiBiMRCxGIhYDEQsBiIWAxGLgYjFQMRiIGIxeAFi7ey97dbN12/gGPzce6+jTHb2tk7gCvXbu6deAUgllp7pGWt9DuP2x6ZjWIPZ/Z/q0eS7joZxBV3tCqUMGbDF2vl7W3tj1i633NIwT67e//zi/6yHXQ0oAkft8m00qA4sikCVwXHW/RtKDbhyAfLKMDrVXNJ4iLTL3ih3GK6ThnmrAvrHuWgwJU/MLKfZMrqz2ei8hAtsFUez2cBSPx9DIcBCyxiiEU7zmw5WO6VzCOPLoLB+NPJQZOo+ocigynoDy2kW1BVoFuDoc88dPbu392PHD41844c3ajZTlx0ewyicDBZmtNpjvBos5cZUEW+pFE7JU8uwVK/lDSNfq+MSTxv56eTBhCUYv36Pjen2ft/e3518n06uDw510BjdOOGXKcZ9GzLCMmkzqqhLdHpD+l0cKccPex1cI67h4atSGSmdf6oYwaWyWNWe5RVmFVZNFIqQZHYfFoh3XCoWIKgUVH9lGpzzRyeQbvz49y/49/BAS7QsViLW3eUZ3AWWBGZ8ACibXA4rpAvFooFygZryjq/v4h95HinHszPnMVaVWXBHmDlYDwilE+0itK8osJBoU8Pzx0gYH6p3t84hjfGBsHVSv12KbCsRK984tSiHclXD0qkVCBKwMavt40IlHz7FsLIWv1daML7yCZ8lcP1ye0w7vp6tVQxk8XPIrpkNdjF8GGJOuc+0+88P+tcvgdWVYQh4KxTac2mffmlkl1mwdxwHn14vLa0A6SllkFEZbgYiFgMRi4GIxUDEYiBiMRCxGKxcLNfY1BYomaXknWozk4zNOGfV83lef9miyHpZuVjTyTVZADcjdUpOE/kB9KFH+YHwURGdVXiH7DurVwN4v082xajYm7Ob10IWZUiGSac3hCzr245lWSQdCqe08pxVGBx0WgByC3c/VD9uP2JyZUAmexappRSAXCoVzWLJPdE+c6yzWrlQxiHaLOv6Fm6OLMTSlQhJMoRcOjyowBmctJUveHgACZXkrFYuVAS/iuh43zmskUwyS+eWY7egCM0jdDUd24YitMyKHhF1VtvlMgqUx8r8H8hILLpjx4EiVJIocxhOXK0SnNVu3UanE7Ou/+Ga4mtELBoGWWXWRiBiMRCxGIhYDEQsBiIWAxGLwcaKFeh5S/xc6Y5J+8FzY8VK04nKRcqQwSKxtJOJLmXQ8KTm5bm2UnrJnLeZdAX/V4GXx0T8UoLDbzeNOqhzETgtBDtRk5fB4umZ5TSrN6faU9GtpKm9TdX3EO75vIuJUFttqH00dpZoJNKJuhyeUYYJ/ZmpvM1oz2c0srB9NDpLBp7qMves6btqSm8z2vMZ0wWa0D46N4sKZeWpLnmDh/yHkknpbfrZ6BKNQHHBBb2D2keDs9BXQdl4qgvFIo+y01N/q+E33DEXcXd5Rj/43mbCFeZ6PlUoHFFdoAntozALCeTNMm6Xp1iq/rxzMG8kiUViqQZO2r6xMXOAW+YC8o0uPbw8bzPpCtPt/WDPJ2zJu+FIsdvYTWgfhVnqdiE4C82rXudGaBYP7o0kIU4pgyXvWZuNiMVAxGIgYjEQsRiIWAxELAYiFgMRi4GIxUDEYiBiMRCxGIhYDEQsBiIWAxGLgYjFQMRiIGIxELEYiFgMRCwGIhYDEYuBiMVAxGIgYjEQsRiIWKkxjH9of8U6yF/dpwAAAABJRU5ErkJggg==";

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

        public static string SplitCamelCase(string input)
        {
            var result = string.Empty;
            if (input != null)
            {
                try
                {
                    result = Regex.Replace(input, "([A-Z])", " $1", RegexOptions.Compiled).Trim();
                }
                catch (Exception)
                {
                    throw;
                }
            }
            return result;
        }

        public static bool IsNullOrWhiteSpace(string value)
        {
            return string.IsNullOrEmpty(value) || value.Trim().Length == 0;
        }

        public static string GetPropertyDisplayName(PropertyInfo pi)
        {
            var dp = pi.GetCustomAttributes(typeof(DisplayNameAttribute), true).Cast<DisplayNameAttribute>().SingleOrDefault();
            return dp != null ? dp.DisplayName : pi.Name;
        }

    }

}
