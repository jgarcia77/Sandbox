namespace Sandbox.Common
{
    using System.Drawing;
    using System.IO;
    using System.Drawing.Imaging;
    using System;

    public enum ImageCompareResults { Equal, Smaller, Bigger, NotDetermined }

    public class ImageHelper
    {
        public string FileName { get; private set; }
        public bool ImageExists { get; private set; }
        public Image Image { get; private set; }
        public string DirectoryName { get; private set; }
        public string FileNameWithoutExtension { get; private set; }
        public string Extension { get; private set; }
        public ImageFormat ImageFormat
        {
            get
            {
                ImageFormat returnValue = null;

                switch (this.Extension)
                { 
                    case ".bmp":
                        returnValue = ImageFormat.Bmp;
                        break;
                    
                    case ".emf":
                        returnValue = ImageFormat.Emf;
                        break;

                    case ".exif":
                        returnValue = ImageFormat.Exif;
                        break;

                    case ".gif":
                        returnValue = ImageFormat.Gif;
                        break;

                    case ".ico":
                        returnValue = ImageFormat.Icon;
                        break;

                    case ".jpeg":
                    case ".jpg":
                        returnValue = ImageFormat.Jpeg;
                        break;

                    case ".dmp":
                        returnValue = ImageFormat.MemoryBmp;
                        break;

                    case ".png":
                        returnValue = ImageFormat.Png;
                        break;

                    case ".tiff":
                        returnValue = ImageFormat.Tiff;
                        break;

                    case ".wmf":
                        returnValue = ImageFormat.Wmf;
                        break;
                }

                return returnValue;
            }
        }

        public string ToExension(ImageFormat imageFormat)
        {
            var returnValue = string.Empty;

            if (imageFormat == ImageFormat.Bmp)
            {
                returnValue = ".bmp";
            }
            else if (imageFormat == ImageFormat.Emf)
            {
                returnValue = ".emf";
            }
            else if (imageFormat == ImageFormat.Exif)
            {
                returnValue = ".exif";
            }
            else if (imageFormat == ImageFormat.Gif)
            {
                returnValue = ".gif";
            }
            else if (imageFormat == ImageFormat.Icon)
            {
                returnValue = ".ico";
            }
            else if (imageFormat == ImageFormat.Jpeg)
            {
                returnValue = ".jpg";
            }
            else if (imageFormat == ImageFormat.MemoryBmp)
            {
                returnValue = ".dmp";
            }
            else if (imageFormat == ImageFormat.Png)
            {
                returnValue = ".png";
            }
            else if (imageFormat == ImageFormat.Tiff)
            {
                returnValue = ".tiff";
            }
            else if (imageFormat == ImageFormat.Wmf)
            {
                returnValue = ".wmf";
            }

            return returnValue;
        }

        public ImageHelper(string fileName)
        {
            this.FileName = fileName;

            if (File.Exists(this.FileName))
            {
                this.Initialize();
                this.ImageExists = true;
            }
            else
            {
                this.ImageExists = false;
            }
        }

        private void Initialize()
        {
            this.Image = Image.FromFile(this.FileName);
            this.DirectoryName = Path.GetDirectoryName(this.FileName);
            this.FileNameWithoutExtension = Path.GetFileNameWithoutExtension(this.FileName);
            this.Extension = Path.GetExtension(this.FileName);

        }

        public ImageCompareResults Compare(Quadrilateral quadrilateral)
        {
            ImageCompareResults returnValue = ImageCompareResults.NotDetermined;

            if (this.Image.Width == quadrilateral.Width && this.Image.Height == quadrilateral.Height)
            {
                returnValue = ImageCompareResults.Equal;
            }
            else if (this.Image.Width < quadrilateral.Width || this.Image.Height < quadrilateral.Height)
            {
                returnValue = ImageCompareResults.Smaller;
            }
            else if (this.Image.Width > quadrilateral.Width || this.Image.Height > quadrilateral.Height)
            {
                returnValue = ImageCompareResults.Bigger;
            }
            
            return returnValue;
        }

        public void ScaleImage(string destination, string fileName, Quadrilateral quadrilateral, ImageFormat imageFormatOverride = null, CroppedDetails croppedDetails = null)
        {
            var ratioX = (double)quadrilateral.Width / this.Image.Width;
            var ratioY = (double)quadrilateral.Height / this.Image.Height;
            var ratio = Math.Min(ratioX, ratioY);

            var newWidth = (int)(this.Image.Width * ratio);
            var newHeight = (int)(this.Image.Height * ratio);

            var newQuadrilateral = new Quadrilateral(newWidth, newHeight);

            this.ResizeImage(destination, fileName, newQuadrilateral, imageFormatOverride, croppedDetails);
        }

        public void ResizeImage(string destination, string fileName, Quadrilateral quadrilateral, ImageFormat imageFormatOverride = null, CroppedDetails croppedDetails = null)
        {
            var imageFromFile = this.Image;

            if (croppedDetails != null)
            {
                var croppedTile = new Bitmap(croppedDetails.Width, croppedDetails.Height);

                croppedTile.SetResolution(this.Image.HorizontalResolution, this.Image.VerticalResolution);

                var croppedGraphic = Graphics.FromImage(croppedTile);

                var croppedArea = new Rectangle(croppedDetails.X1, croppedDetails.Y1, croppedDetails.Width, croppedDetails.Height);

                croppedGraphic.DrawImage(this.Image, 0, 0, croppedArea, GraphicsUnit.Pixel);

                imageFromFile = croppedTile;
            }

            var newImage = new Bitmap(quadrilateral.Width, quadrilateral.Height);

            newImage.SetResolution(imageFromFile.HorizontalResolution, imageFromFile.VerticalResolution);
                        
            Graphics.FromImage(newImage)
                    .DrawImage(imageFromFile, 0, 0, quadrilateral.Width, quadrilateral.Height);

            if (imageFormatOverride == null)
            {
                var fileNameNew = string.Concat(destination, "//", fileName, this.Extension);
                newImage.Save(fileNameNew, this.ImageFormat);
            }
            else
            {
                var extension = this.ToExension(imageFormatOverride);
                var fileNameNew = string.Concat(destination, "//", fileName, extension);
                newImage.Save(fileNameNew, imageFormatOverride);
            }
        }
    }

    public struct Quadrilateral
    {
        public int Width;
        public int Height;

        public Quadrilateral(int width, int height)
        {
            Width = width;
            Height = height;
        }
    }

    public class CroppedDetails
    {
        public int X1 { get; private set; }
        public int Y1 { get; private set; }
        public int X2 { get; private set; }
        public int Y2 { get; private set; }
        public int Width { get; private set; }
        public int Height { get; private set; }

        public CroppedDetails(int x1, int y1, int x2, int y2, int width, int height)
        {
            this.X1 = x1;
            this.Y1 = y1;
            this.X2 = x2;
            this.Y2 = y2;
            this.Width = width;
            this.Height = height;
        }
    }
}
