using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.IO;

namespace CNO.BPA.CreateMDWCover
{
    class TIFFBuilder
    {

        public string createTIFF(string Contents)
        {
            string fileName = string.Empty;
            Bitmap source = null;
            Int32 textLocation = 100;

            fileName = GetTempFilePathWithExtension("tif");

            source = new Bitmap(2550, 3300);
            source.SetResolution(300, 300);
            
            Font objFont = new Font("Arial", 40, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Pixel);

            // Create a graphics object to measure the text's width and height.
            Graphics objGraphics = Graphics.FromImage(source);
            StringFormat textFormat = new StringFormat();
          
            objGraphics.Clear(Color.White);
            objGraphics.DrawString(Contents, objFont, new SolidBrush(Color.Black),60, textLocation);
            objGraphics.Flush();
            
            EncoderParameters ep = new EncoderParameters(2);
            ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.SaveFlag, (long)EncoderValue.MultiFrame);
            ep.Param[1] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)System.Drawing.Imaging.EncoderValue.CompressionCCITT4);

            ImageCodecInfo info = GetEncoderInfo("image/tiff");

            source = ConvertToBitonal(source);

            source.Save(fileName, info, ep);
            source.Dispose();
            objGraphics.Dispose();
            return fileName;

        }
        public static Bitmap ConvertToBitonal(Bitmap original)
        {
            try
            {

                Bitmap source = null;

                // If original bitmap is not already in 32 BPP, ARGB format, then convert
                if (original.PixelFormat != PixelFormat.Format32bppArgb)
                {
                    source = new Bitmap(original.Width, original.Height, PixelFormat.Format32bppArgb);
                    source.SetResolution(original.HorizontalResolution, original.VerticalResolution);
                    using (Graphics g = Graphics.FromImage(source))
                    {
                        g.DrawImageUnscaled(original, 0, 0);
                    }
                }
                else
                {
                    source = new Bitmap(original);
                    source.SetResolution(original.HorizontalResolution, original.VerticalResolution);

                }

                // Lock source bitmap in memory
                BitmapData sourceData = source.LockBits(new Rectangle(0, 0, source.Width, source.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

                // Copy image data to binary array
                int imageSize = sourceData.Stride * sourceData.Height;
                byte[] sourceBuffer = new byte[imageSize];
                Marshal.Copy(sourceData.Scan0, sourceBuffer, 0, imageSize);

                // Unlock source bitmap
                source.UnlockBits(sourceData);

                // Create destination bitmap
                Bitmap destination = new Bitmap(source.Width, source.Height, PixelFormat.Format1bppIndexed);
                destination.SetResolution(original.HorizontalResolution, original.VerticalResolution);

                // Lock destination bitmap in memory
                BitmapData destinationData = destination.LockBits(new Rectangle(0, 0, destination.Width, destination.Height), ImageLockMode.WriteOnly, PixelFormat.Format1bppIndexed);

                // Create destination buffer
                imageSize = destinationData.Stride * destinationData.Height;
                byte[] destinationBuffer = new byte[imageSize];

                int sourceIndex = 0;
                int destinationIndex = 0;
                int pixelTotal = 0;
                byte destinationValue = 0;
                int pixelValue = 128;
                int height = source.Height;
                int width = source.Width;
                int threshold = 500;

                // Iterate lines
                for (int y = 0; y < height; y++)
                {
                    sourceIndex = y * sourceData.Stride;
                    destinationIndex = y * destinationData.Stride;
                    destinationValue = 0;
                    pixelValue = 128;

                    // Iterate pixels
                    for (int x = 0; x < width; x++)
                    {
                        // Compute pixel brightness (i.e. total of Red, Green, and Blue values)
                        pixelTotal = sourceBuffer[sourceIndex + 1] + sourceBuffer[sourceIndex + 2] + sourceBuffer[sourceIndex + 3];
                        if (pixelTotal > threshold)
                        {
                            destinationValue += (byte)pixelValue;
                        }
                        if (pixelValue == 1)
                        {
                            destinationBuffer[destinationIndex] = destinationValue;
                            destinationIndex++;
                            destinationValue = 0;
                            pixelValue = 128;
                        }
                        else
                        {
                            pixelValue >>= 1;
                        }
                        sourceIndex += 4;
                    }
                    if (pixelValue != 128)
                    {
                        destinationBuffer[destinationIndex] = destinationValue;
                    }
                }

                // Copy binary image data to destination bitmap
                Marshal.Copy(destinationBuffer, 0, destinationData.Scan0, imageSize);

                // Unlock destination bitmap
                destination.UnlockBits(destinationData);

                // Return
                return destination;

            }
            catch (Exception ex)
            {
                throw new Exception("TiffUtility.ConvertToBitonal: " + ex.Message);
            }


        }
        private ImageCodecInfo GetEncoderInfo(string mimeType)
        {

            ImageCodecInfo[] encoders = ImageCodecInfo.GetImageEncoders();
            for (int j = 0; j < encoders.Length; j++)
            {
                if (encoders[j].MimeType == mimeType)
                    return encoders[j];
            }

            throw new Exception(mimeType + " mime type not found in ImageCodecInfo");
        }
        private static string GetTempFilePathWithExtension(string extension) 
        { 
            string path = Path.GetTempPath();
            string fileName = Guid.NewGuid().ToString();
            string tempfile = Path.Combine(path, fileName);
            return Path.ChangeExtension(tempfile, extension);
        } 

    }
}
