using BarcodeLib.Barcode;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LPF.Printer
{
    public class BarcodeInfo
    {
        #region[创建一个一维码]
        /// <summary>
        /// 创建一个一维码
        /// </summary>
        /// <param name="text">一维码内容</param>
        /// <param name="filePath">完整存放路径（包含文件名）</param>
        /// <param name="width">图片宽，不传则原始图片默认大小</param>
        /// <param name="height">图片高，不传则原始图片默认大小</param>
        public bool CreateBarCode(string text, string filePath, int? width = null, int? height = null)
        {
            bool temp = true;
            try
            {
                if (width > 0 && height > 0)
                {
                    Linear code128 = CreateBarCodeBySize(text.Trim(), width, height);
                    code128.drawBarcode(filePath);
                }
                else
                {
                    BarcodeLib.Barcode.Linear code128 = CreateBarCode(text.Trim());
                    code128.drawBarcode(filePath);
                }
            }
            catch
            {
                temp = false;
            }
            return temp;
        }

                /// <summary>
        /// 创建一个一维码
        /// </summary>
        /// <param name="text">一维码内容</param>
        /// <param name="filePath">完整存放路径（包含文件名）</param>
        /// <param name="width">图片宽，不传则原始图片默认大小</param>
        /// <param name="height">图片高，不传则原始图片默认大小</param>
        private Linear CreateBarCode(string text)
        {
            BarcodeLib.Barcode.Linear code128 = new BarcodeLib.Barcode.Linear();
            code128.Type = BarcodeType.CODE128;
            code128.Data = text;
            code128.AddCheckSum = true;
            code128.UOM = UnitOfMeasure.PIXEL;
            code128.ShowText = false;
            code128.BarWidth = 1;
            code128.BarHeight = 1;
            code128.LeftMargin = 0;
            code128.RightMargin = 0;
            code128.ImageFormat = System.Drawing.Imaging.ImageFormat.Bmp;
            return code128;
        }

        /// <summary>
        /// 创建一个一维码
        /// </summary>
        /// <param name="text">一维码内容</param>
        /// <param name="width">图片宽，不传则原始图片默认大小</param>
        /// <param name="height">图片高，不传则原始图片默认大小</param>
        private Linear CreateBarCodeBySize(string text, int? width = null, int? height = null)
        {
            BarcodeLib.Barcode.Linear code128 = new BarcodeLib.Barcode.Linear();
            code128.Type = BarcodeType.CODE128;
            code128.Data = text;
            code128.AddCheckSum = true;
            code128.UOM = UnitOfMeasure.PIXEL;
            code128.ShowText = false;
            if (width == null || width < 0)
            {
                width = 1;
            }
            code128.BarWidth = (float)width;
            if (height == null || height < 0)
            {
                height = 1;
            }
            code128.BarHeight = (float)height;
            code128.LeftMargin = 0;
            code128.RightMargin = 0;
            code128.ImageFormat = System.Drawing.Imaging.ImageFormat.Bmp;
            return code128;
        }
        #endregion
    }
}
