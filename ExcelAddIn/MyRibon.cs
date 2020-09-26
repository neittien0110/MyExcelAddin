using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Diagnostics;

namespace ExcelAddIn
{
    public partial class MyRibon
    {
        private void MyRibon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonImage2Cells_Click(object sender, RibbonControlEventArgs e)
        {
            Bitmap img;
            const int MAX_HEIGHT = 320;
            const int MAX_PIXEL = 82455; //chính xác đúng ngần này điểm

            /// Tạo dialog để chọn file ảnh
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                /// Chỉ chấp nhận file dạnh ảnh
                dialog.Filter = "image files (*.jpg)|*.jpg|*.png|*.png|*.bmp|*.bmp|All files (*.*)|*.*";
                dialog.FilterIndex = 1;

                if (dialog.ShowDialog() != DialogResult.OK)
                {
                    return;
                }

                /// Mở file ảnh
                img = new Bitmap(dialog.FileName);

                /// Tự co giãn tỷ lệ theo số điểm tối đa            
                double ratio;
                int newwwidth = img.Width;
                int newheight = img.Height;
                if ( img.Width * img.Height > MAX_PIXEL)
                {
                    ratio = Math.Sqrt((double)MAX_PIXEL / img.Width / img.Height);
                    newwwidth = (int)(img.Width * ratio);
                    newheight = (int)(img.Height * ratio);
                    img = ResizeBitmap(img, newwwidth, newheight);
                }

                /*
                /// Tự co giãn tỷ lệ theo chiều cao và chiều dọc để không vượt qua
                ratio = img.Width / img.Height;
                if (newheight > MAX_HEIGHT)
                {
                    newheight = MAX_HEIGHT;
                    newwwidth = (int)(newheight * ratio);
                }
                if (newwwidth > MAX_HEIGHT)
                {
                    newwwidth = MAX_HEIGHT;
                    newheight = (int)(newwwidth / ratio);
                }
                img = ResizeBitmap(img, newwwidth, newheight);
                */



                dialog.Dispose();
            }



            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (wb == null)
            {
                MessageBox.Show("Bạn phải mở một workbook.");
                img.Dispose();
                return;
            }

            /// Lấy sheet đang được active hiện thời
            Worksheet ws = wb.ActiveSheet;

            Task.Run(() =>
            {
                /// Đưa chiều rộng của các cột bằng nhau và bằng 2
                for (int i = 1; i <= img.Width; i++)
                {
                    ws.Columns[i].ColumnWidth = 2;
                }
            });

            Task.Run(() =>
                {
                    /// Biến đếm số điểm ảnh 
                    int count = 0;
                    /// Đặt cờ báo hiệu có lỗi trong quá trình convert
                    bool error_flag = false;
                    /// Đọc từng picel ảnh và qui đổi thành màu nền của cell
                    for (int i = 0; i < img.Height; i++)
                    {
                        for (int j = 0; j < img.Width; j++)
                        {
                            Color pixel = img.GetPixel(j, i);
                            retry:
                            try
                            {
                                ws.Cells[i + 1, j + 1].Interior.Color = pixel.R | (pixel.G << 8) | (pixel.B << 16);
                            }
                            catch (Exception ex)  //user click chuột vào 1 cell là sinh ngoại lệ và dừng ngay.
                            {
                                if ((uint)ex.HResult == 0x800a03ec)
                                {
                                    error_flag = true; 
                                    goto _end_of_image;
                                }
                                else
                                {
                                    Debug.WriteLine(ex.Message);
                                    Debug.WriteLine("i={0}, j={1}", i, j);
                                    Task.Delay(250);
                                    goto retry;
                                }
                            }
                            count++;
                            if (count == 82455)
                            {
                                int x;
                                x = count;
                            }
                        }

                    }
                    _end_of_image:
                    if (error_flag)
                    {
                        MessageBox.Show("Excel cannot process too many different cell formats. Please create another workbook.", "Error 0x800a03ec");
                    }
                    else
                    {
                        MessageBox.Show("Finish converting from image " + img.Height + "x" + img.Width + " pixels to " + count + " cells. Have fun!");
                    }
                    img.Dispose();
                }
            );

        }
        /// <summary>
        ///     Zoom ảnh
        /// </summary>
        /// <param name="bmp">Đối tượng cần zoom </param>
        /// <param name="width">chiều ngang mong muốn</param>
        /// <param name="height">chiều dọc mong muốn</param>
        /// <returns></returns>
        static Bitmap ResizeBitmap(Bitmap bmp, int width, int height)
        {
            Bitmap result = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(result))
            {
                g.DrawImage(bmp, 0, 0, width, height);
            }

            return result;
        }
    }
}
