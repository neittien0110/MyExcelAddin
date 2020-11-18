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
    /// <remarks>
    ///         Có 2 họ hàm
    ///             + Sử dụng Google API service, yêu cầu phải có kết nối mạng khi sử dụng
    ///             + Sử dụng thư viện ZXing, không cần có mạng
    /// </remarks>
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

        /// <summary>
        ///         Trả về file name (path) của ảnh tạo ra bởi ZXing library (temp file)
        /// </summary>
        /// <param name="Text">Văn bản cần sinh mã QR</param>
        /// <param name="ImageSize">Kích thước của ảnh QR. Tối đa là 500 px</param>
        /// <param name="Correction">Mức độ chịu lỗi</param>
        /// <param name="Margin">Số điểm ảnh trắng để làm biên </param>
        /// <returns></returns>
        static string GetQRCodeLocalFileNameByZXing(string Text, int ImageSize = 500, CorrectionLevel Correction = CorrectionLevel.High, int Margin = 0)
        {
            QRCodeWriter qr = new ZXing.QrCode.QRCodeWriter(); //QRCode as a BitMatrix 2D array

            Dictionary<EncodeHintType, object> hint = new Dictionary<EncodeHintType, object>();
            hint.Add(EncodeHintType.MARGIN, Margin); // margin of the QRCode image
            hint.Add(EncodeHintType.ERROR_CORRECTION, Correction);

            var matrix = qr.encode(Text, BarcodeFormat.QR_CODE, ImageSize, ImageSize, hint); // encode QRCode matrix from source text
            ZXing.BarcodeWriter w = new ZXing.BarcodeWriter();
            Bitmap img = w.Write(matrix); // QRCode Bitmap image
            string tempFile = Path.GetTempFileName(); //create a temp file to save QRCode image
            img.Save(tempFile, System.Drawing.Imaging.ImageFormat.Png);//save QRCode image to temp file

            return tempFile;
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
                    int option = 1; // luôn sử dụng thư viện ZXing
                    if (option == 1)
                    {
                        MyShape.Fill.UserPicture(Barcode.GetQRCodeLocalFileNameByZXing(Text, 500, CorrectionLevel.High, Margin));
                    } 
                    else if (option == 2)
                    {
                        MyShape.Fill.UserPicture(Barcode.GetQRCodeWebAPI(Text, 500, CorrectionLevel.High, Margin));
                    }
                    
                }
                catch
                {
                    return "Disconnect";
                }
            }
            return Text;
        }

        // Old codes before refactor
        //Generate QRCode using ZXing library
        //[ExcelDna.Integration.ExcelFunction(Description = "QRCode generator by ZXing lib")]
        //public static object QRCodeZ(
        //    [ExcelDna.Integration.ExcelArgument(Description = "Shape name to contain QRCode image (view Selection Pane). If not existed, new shape will be created. Example: tl123")]
        //    string ShapeName,
        //    [ExcelDna.Integration.ExcelArgument(Description = "Text to be transformed to QRCode. Example: \"hello\"")] 
        //    string Text)
        //{
        //    Application xlApp = (Application)ExcelDnaUtil.Application;
        //    Workbook wb = xlApp.ActiveWorkbook;
        //    if (wb == null) return "";
        //    Worksheet ws = wb.ActiveSheet;

        //    QRCodeWriter qr = new ZXing.QrCode.QRCodeWriter(); //QRCode as a BitMatrix 2D array

        //    Dictionary<EncodeHintType, object> hint = new Dictionary<EncodeHintType, object>();
        //    hint.Add(EncodeHintType.MARGIN, 0); // margin of the QRCode image


        //    var matrix = qr.encode(Text, BarcodeFormat.QR_CODE, 50, 50, hint); // encode QRCode matrix from source text
            
        //    ZXing.BarcodeWriter w = new ZXing.BarcodeWriter();
        //    Bitmap img = w.Write(matrix); // QRCode Bitmap image
        //    string tempFile = Path.GetTempFileName();

        //    img.Save(tempFile, System.Drawing.Imaging.ImageFormat.Png);


        //    Shape MyShape = null;

           
        //    MyShape = ws.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, xlApp.ActiveCell.Left, xlApp.ActiveCell.Top, xlApp.ActiveCell.Width, xlApp.ActiveCell.Height);
            
            
        //    MyShape.Fill.UserPicture(tempFile);
            
            
        //    return Text;
        //}
    }
}
