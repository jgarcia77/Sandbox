using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;

namespace Sandbox.ImageResizer
{
    class Program
    {
        static void Main(string[] args)
        {
            var file = @"C:\Users\JosueG\Desktop\Images\jordan.jpeg";

            ResizeImage(file, 128, 128);
            ResizeImage(file, 64, 64);
            ResizeImage(file, 32, 32);
            ResizeImage(file, 16, 16);
        }

        public static void ResizeImage(string file, int maxWidth, int maxHeight)
        {
            var fullPath = Path.GetDirectoryName(file);
            var fileName = Path.GetFileNameWithoutExtension(file);
            var fileExtension = Path.GetExtension(file);
            var imageFormat = GetImageFormat(fileExtension);

            var imageFromFile = Image.FromFile(file);
            var newImage = ScaleImage(imageFromFile, maxHeight, maxWidth);

            var fileSuffix = string.Concat("-", maxWidth, "x", maxHeight);
            var fileNameNew = string.Concat(fullPath, "//", fileName, fileSuffix, fileExtension);
            newImage.Save(fileNameNew, imageFormat);
        }

        public static Image ScaleImage(Image image, int maxWidth, int maxHeight)
        {
            var ratioX = (double)maxWidth / image.Width;
            var ratioY = (double)maxHeight / image.Height;
            var ratio = Math.Min(ratioX, ratioY);

            var newWidth = (int)(image.Width * ratio);
            var newHeight = (int)(image.Height * ratio);

            var newImage = new Bitmap(newWidth, newHeight);
            Graphics.FromImage(newImage).DrawImage(image, 0, 0, newWidth, newHeight);
            return newImage;
        }

        private static ImageFormat GetImageFormat(string extension)
        {
            ImageFormat returnValue = null;

            switch (extension)
            { 
                case ".jpeg":
                    returnValue = ImageFormat.Jpeg;
                    break;

            }

            return returnValue;
        }

    }
}
