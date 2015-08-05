using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;
using Sandbox.Common;

namespace Sandbox.ImageResizer
{
    class Program
    {
        static void Main(string[] args)
        {
            //ImageHelperConstructorTestFail();
            //ImageHelperConstructorTestPass();
            //ImageHelper_CompareResults();
            //ImageHelper_Resize();
            //ImageHelper_Resize_Shrink();
            //ImageHelper_Resize_Shrink_ToPng();
            ImageHelper_Resize_Crop_ToPng();

            var emptyGuid = Guid.Empty.ToString();

            Console.Read();
        }

        private static void ImageHelperConstructorTestFail()
        {
            var fileName = @"C:\Users\JosueG\Desktop\Images\jordan.jpeg";
            var helper = new ImageHelper(fileName);

            Console.WriteLine(helper.ImageExists);
        }

        private static void ImageHelperConstructorTestPass()
        {
            var fileName = @"C:\ProgramData\DirectiveBoards\Uploads\Images\Temp\logo-light.png";
            var helper = new ImageHelper(fileName);
            
            Console.WriteLine(helper.ImageExists);

            var dimensions = "Width = {0}; Height = {1}";

            Console.WriteLine(string.Format(dimensions, helper.Image.Width, helper.Image.Height));
        }

        private static void ImageHelper_CompareResults()
        {
            var fileName = @"C:\ProgramData\DirectiveBoards\Uploads\Images\Temp\avatar.png";
            var helper = new ImageHelper(fileName);

            var quadrilateral = new Quadrilateral(45, 45);
            var result = helper.Compare(quadrilateral);

            OutputImageCompare(helper.Image, quadrilateral, result);

            quadrilateral = new Quadrilateral(90, 25);
            result = helper.Compare(quadrilateral);

            OutputImageCompare(helper.Image, quadrilateral, result);

            quadrilateral = new Quadrilateral(25, 90);
            result = helper.Compare(quadrilateral);

            OutputImageCompare(helper.Image, quadrilateral, result);

            quadrilateral = new Quadrilateral(25, 90);
            result = helper.Compare(quadrilateral);

            OutputImageCompare(helper.Image, quadrilateral, result);

            quadrilateral = new Quadrilateral(25, 25);
            result = helper.Compare(quadrilateral);

            OutputImageCompare(helper.Image, quadrilateral, result);
        }

        private static void OutputImageCompare(Image image, Quadrilateral quadrilateral, ImageCompareResults result)
        {
            Console.WriteLine("--------------------------------------------------");
            Console.WriteLine("Result: {0}", result);
            Console.WriteLine("Image Width: {0}", image.Width);
            Console.WriteLine("Image Height: {0}", image.Height);
            Console.WriteLine("New Width: {0}", quadrilateral.Width);
            Console.WriteLine("New Height: {0}", quadrilateral.Height);
        }

        private static void ImageHelper_Resize()
        {
            var fileName = @"C:\ProgramData\DirectiveBoards\Uploads\Images\Temp\logo-light cowboy1.png";
            var helper = new ImageHelper(fileName);

            var quadrilateral = new Quadrilateral(106, 17);
            helper.ResizeImage(helper.DirectoryName, "lense_xs", quadrilateral, ImageFormat.Png);

            //quadrilateral = new Quadrilateral(45, 45);
            //helper.ResizeImage(helper.DirectoryName, "lense_sm", quadrilateral, ImageFormat.Png);

            //quadrilateral = new Quadrilateral(150, 150);
            //helper.ResizeImage(helper.DirectoryName, "lense_md", quadrilateral, ImageFormat.Png);

            //quadrilateral = new Quadrilateral(300, 300);
            //helper.ResizeImage(helper.DirectoryName, "lense_lg", quadrilateral, ImageFormat.Png);

            //quadrilateral = new Quadrilateral(600, 600);
            //helper.ResizeImage(helper.DirectoryName, "lense_xl", quadrilateral, ImageFormat.Png);
            
        }

        private static void ImageHelper_Resize_Shrink()
        {
            var fileName = @"C:\ProgramData\DirectiveBoards\Uploads\Images\Temp\oberon.jpg";
            var helper = new ImageHelper(fileName);

            var quadrilateral = new Quadrilateral(45, 45);
            helper.ScaleImage(helper.DirectoryName, "oberon_s45_45", quadrilateral);
            helper.ResizeImage(helper.DirectoryName, "oberon_r45_45", quadrilateral);
        }

        private static void ImageHelper_Resize_Shrink_ToPng()
        {
            var fileName = @"C:\ProgramData\DirectiveBoards\Uploads\Images\Temp\oberon.jpg";
            var helper = new ImageHelper(fileName);

            var quadrilateral = new Quadrilateral(45, 45);
            helper.ScaleImage(helper.DirectoryName, "oberon_s45_45", quadrilateral, ImageFormat.Png);
            helper.ResizeImage(helper.DirectoryName, "oberon_r45_45", quadrilateral, ImageFormat.Png);
        }

        private static void ImageHelper_Resize_Crop_ToPng()
        {
            var fileName = @"C:\ProgramData\DirectiveBoards\Uploads\Images\Temp\dpavatar1.jpg";
            var helper = new ImageHelper(fileName);

            var quadrilateral = new Quadrilateral(16, 16);
            var croppedDetails = new CroppedDetails(64, 40, 202, 151, 138, 111);

            helper.ResizeImage(helper.DirectoryName, "icon_16_16", quadrilateral, ImageFormat.Png, croppedDetails);
            helper.ResizeImage(helper.DirectoryName, "icon_16_16", quadrilateral, ImageFormat.Icon, croppedDetails);
        }

        public static void ResizeImage()
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
