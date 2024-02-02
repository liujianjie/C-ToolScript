using System;
using System.IO;
using Aspose.Cells;
using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Cells.Drawing;

// 压缩Excel图片大小
namespace CompressExcelPictureSize.MainFloder
{
    internal class ExcelImageCopressor
    {
        static void Main()
        {
            // 读取 Excel 文件
            string filePath = "Baseball_ScreenShot_WholeKey.xlsx";
            Workbook workbook = new Workbook(filePath);

            // 获取第一个工作表
            Worksheet worksheet = workbook.Worksheets[0];

            // 读取图片
            PictureCollection pictures = worksheet.Pictures;
            int compressionQuality = 75; // 设置压缩质量 (0-100，值越小压缩率越高，图片质量越低)

            for (int i = 0; i < pictures.Count; i++)
            {
                Picture picture = pictures[i];

                // 将图片数据转换为 System.Drawing.Bitmap
                MemoryStream imageStream = new MemoryStream(picture.Data);
                Bitmap originalImage = new Bitmap(imageStream);

                // 使用 System.Drawing.Common 进行图片压缩
                byte[] compressedImageData = CompressImage(originalImage, compressionQuality);

                // 更新图片数据
                picture.Data = compressedImageData;
            }

            // 保存修改后的 Excel 文件
            string outputPath = "file1.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }

        static byte[] CompressImage(Bitmap originalImage, int compressionQuality)
        {
            using (MemoryStream compressedImageStream = new MemoryStream())
            {
                EncoderParameters encoderParameters = new EncoderParameters(1);
                encoderParameters.Param[0] = new EncoderParameter(Encoder.Quality, compressionQuality);

                ImageCodecInfo jpegCodec = GetEncoderInfo("image/jpeg");
                originalImage.Save(compressedImageStream, jpegCodec, encoderParameters);

                return compressedImageStream.ToArray();
            }
        }

        static ImageCodecInfo GetEncoderInfo(string mimeType)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageEncoders();
            for (int i = 0; i < codecs.Length; i++)
            {
                if (codecs[i].MimeType == mimeType)
                {
                    return codecs[i];
                }
            }
            return null;
        }
    }
}
