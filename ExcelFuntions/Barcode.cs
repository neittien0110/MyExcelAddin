using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;       // Cài đặt Microsoft.Office.Core (Nuget) và Add Reference Microsoft.Office.Interop.Excel
using ZXing;
using ZXing.Common;
using ZXing.QrCode;

namespace MyExcelAddIn
{
    /// <summary>
    ///         Các hàm Excel liên quan tới sinh mã barcode
    /// </summary>
    public class Barcode
    {
        /// <summary>
        ///       Mức độ hiệu chỉnh lỗi của mã QRCode
        /// </summary>
        public enum CorrectionLevel
        {
            /// <summary> Cho phép hỏng 7% ảnh barcode  </summary>
            Low = 'L',
            /// <summary> Cho phép hỏng 15% ảnh barcode  </summary>
            Medium = 'M',
            /// <summary> Cho phép hỏng 25% ảnh barcode  </summary>
            Quad = 'Q',
            /// <summary> Cho phép hỏng 30% ảnh barcode  </summary>
            High = 'H'
        }
            
        /// <summary>
        ///         Trả về link ảnh từ dịch vụ QRCode của Google
        /// </summary>
        /// <param name="Text">Văn bản cần sinh mã QR</param>
        /// <param name="ImageSize">Kích thước của ảnh QR. Tối đa là 500 px</param>
        /// <param name="Correction">Mức độ chịu lỗi</param>
        /// <param name="Margin">Số điểm ảnh trắng để làm biên </param>
        /// <returns></returns>
        static string GetQRCodeWebAPI(string Text, int ImageSize = 500, CorrectionLevel Correction = CorrectionLevel.High, int Margin = 0)
        {
            StringBuilder sURL = new StringBuilder();
            sURL.AppendFormat("https://chart.googleapis.com/chart?cht=qr&chs={0}x{0}&chld={1}|{2}&chl={3}", ImageSize, Correction, Margin, Text);
            return sURL.ToString();
        }

        [ExcelDna.Integration.ExcelFunction(Description = "Tạo mã QRCode")]
        public static object QRCode(
            [ExcelDna.Integration.ExcelArgument(Description = "Tên của Shape sẽ chứa ảnh QRCode (xem bằng Selection Pane). Nếu shape chưa tồn tại, hàm sẽ tự tạo mới. Ví dụ: tl123")] 
            string ShapeName,
            [ExcelDna.Integration.ExcelArgument(Description = "Văn bản cần chuyển thành QRcode. Ví dụ: xin chào bạn")]
            string Text,
            [ExcelDna.Integration.ExcelArgument(Description = "Khoảng trắng giữa QRCode và viên ngoài của Shape. Ví dụ: 0", Name="[Margin=0]")]
            int Margin = 0)
        {
            Application xlApp = (Application)ExcelDnaUtil.Application;

            Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return "";

            Worksheet ws = wb.ActiveSheet;

            Shape MyShape = null;

            /// Tìm xem có Shape nào có tên như tham số vào không
            foreach (Shape shape in ws.Shapes)
                if (shape.Name == ShapeName)
                {
                    MyShape = shape;
                }

            /// Nếu chưa có Shape thì tự tạo  mới luôn
            if (MyShape == null)
            {
                MyShape = ws.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, xlApp.ActiveCell.Left, xlApp.ActiveCell.Top, xlApp.ActiveCell.Width, xlApp.ActiveCell.Height);
                MyShape.Name = ShapeName;
                MyShape.Line.Transparency = (float)(1.0);
                MyShape.Fill.Solid();
                MyShape.Fill.ForeColor.RGB = 0xEEEEEE;
            };
            /// Và đặt vào đó hình ảnh QRCode
            {
                try
                {
                    MyShape.Fill.UserPicture(Barcode.GetQRCodeWebAPI(Text, 500, CorrectionLevel.High, Margin));
                }
                catch
                {
                    return "Disconnect";
                }
            }
            return Text;
        }

        //Generate QRCode using ZXing library
        [ExcelDna.Integration.ExcelFunction(Description = "QRCode generator by ZXing lib")]
        public static object QRCodeZ(
            [ExcelDna.Integration.ExcelArgument(Description = "Text to be transformed to QRCode. Example: \"hello\"")] 
            string Text)
        {
            string imageFile = "C:\\Users\\ASUS\\Desktop\\myQR.png"; // image file to save as QRCode image, change path in different devices
            QRCodeWriter qr = new ZXing.QrCode.QRCodeWriter(); //QRCode as a BitMatrix 2D array

            Dictionary<EncodeHintType, object> hint = new Dictionary<EncodeHintType, object>();
            hint.Add(EncodeHintType.MARGIN, 0); // margin of the QRCode image


            var matrix = qr.encode(Text, BarcodeFormat.QR_CODE, 350, 350, hint); // encode QRCode matrix from source text
            
            ZXing.BarcodeWriter w = new ZXing.BarcodeWriter();
            Bitmap img = w.Write(matrix); // QRCode Bitmap image
            //img.Save(@imageFile, System.Drawing.Imaging.ImageFormat.Png); // save QRCode image to local device
            
            MemoryStream ms = new MemoryStream(); 
            img.Save(ms, System.Drawing.Imaging.ImageFormat.Png); //save QRCode image to memory stream
            System.Drawing.Image i = System.Drawing.Image.FromStream(ms); // create image in Image form
            
            Application xlApp = (Application)ExcelDnaUtil.Application;
            Workbook wb = xlApp.ActiveWorkbook;
            if (wb == null) return "";
            Worksheet ws = wb.ActiveSheet;

            //the following codes took me hours, just leave them here for references
            //add QRCode image to this worksheet
            //ws.Shapes.AddPicture(@imageFile, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 0, 185, 185);
            //Shape myShape = ws.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, xlApp.ActiveCell.Left, xlApp.ActiveCell.Top, xlApp.ActiveCell.Width, xlApp.ActiveCell.Height);
            //System.Drawing.Image.FromStream(ms);
            //myShape.Fill.UserPicture(imageMemoryStream.ToString());
            //ws.
            //ws.Shapes.AddPicture(i, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 0, 0, 185, 185);
            
            Clipboard.SetDataObject(i); //set image to clipboard
            ws.Paste(); //just paste the clipboard
            
            return Text;
        }
    }
}
