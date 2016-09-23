using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Web;

namespace Word2Html
{
    public class ImageHelper
    {
        /// <summary>
        /// 图片转Base64编号
        /// </summary>
        /// <param name="imageFile"></param>
        /// <returns></returns>
        public static string ImageToBase64(string imageFile)
        {
            MemoryStream m = new MemoryStream();
            Bitmap bp = new Bitmap(imageFile);
            bp.Save(m, ImageFormat.Png);
            byte[] b = m.GetBuffer();
            string base64String = Convert.ToBase64String(b);
            return base64String;
        }

        /// <summary>
        /// Base64编号转为图片
        /// </summary>
        /// <param name="base64String"></param>
        /// <param name="fileName"></param>
        /// <param name="format"></param>
        public static void Base64ToImage(string base64String, string fileName, ImageFormat format)
        {
            byte[] bt = Convert.FromBase64String(base64String);
            MemoryStream stream = new MemoryStream(bt);
            Bitmap bitmap = new Bitmap(stream);
            bitmap.Save(fileName, format);
        }
    }
}