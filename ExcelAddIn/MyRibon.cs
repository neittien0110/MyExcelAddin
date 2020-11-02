using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;
using System.Drawing;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Speech.Recognition;

namespace ExcelAddIn
{
    public partial class MyRibon
    {
        private void MyRibon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void buttonImage2Cells_Click(object sender, RibbonControlEventArgs e)
        {   //load image into cells - by MSc Tien
            Bitmap img;
            //const int MAX_HEIGHT = 320;
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
        /// 

        private void buttonColorize_Click(string color, string saturation)
        {   //Colorize the cells based on selected color and saturation - by Stnd Tuong

            //get selected cells
            Range currentRange = (Range)Globals.ThisAddIn.Application.Selection as
                Microsoft.Office.Interop.Excel.Range;
            if (currentRange == null) return;

            //Read each cell and colorize it
            //foreach(var mycell in currentRange.Cells)
            //{
            //    //ignore null cells
            //    if(mycell!= null && ((dynamic)(mycell)).Value != null)
            //    {
            //        currentRange.Interior.Color = 37;
            //    }
            //}

            
            int saturationInt; //saturation value in integer 
            bool isNumeric = int.TryParse(saturation,out saturationInt); //boolean to check if saturation is a valid number
            if (isNumeric) //check if the saturation is a number
            {
                if (saturationInt >= 0 && saturationInt <= 255) //check if the saturation value is in range 0 to 255
                {
                    switch (color) // color to display with saturation 
                    {
                        case "AppointmentColor1": currentRange.Interior.Color = Color.FromArgb(saturationInt, 0, 0); break; //red
                        case "AppointmentColor2": currentRange.Interior.Color = Color.FromArgb(0, 0, saturationInt); break; //green
                        case "AppointmentColor3": currentRange.Interior.Color = Color.FromArgb(0, saturationInt, 0); break; //blue
                        case "AppointmentColor4": currentRange.Interior.Color = Color.FromArgb(saturationInt, saturationInt, saturationInt); break; //gray
                        default: break;
                    }
                    if (saturationInt > 128) // change font color to black or white based on background's saturation
                    {
                        currentRange.Font.Color = Color.Black;
                    }
                    else
                    {
                        currentRange.Font.Color = Color.White;
                    }
                    currentRange.Value = saturationInt; // cell value = saturation value
                }
                else
                {
                    //saturation is not in range 0 to 255
                    MessageBox.Show("Please input a number between 0 and 255 at saturation box", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                //saturation is not a valid number
                MessageBox.Show("Invalid type, please input a number at saturation box", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
           
            
            //test code:
            //currentRange.Value = color;
            //currentRange.Interior.Color = Color.FromArgb(0, 200, 0);

        }

        private static SpeechRecognitionEngine engine;
        //private System.Windows.Forms.TextBox textBox;
        private void buttonCortana_Click()
        {   //Speech recognition

            //get selected cells
            //Range currentRange = (Range)Globals.ThisAddIn.Application.Selection as
            //    Microsoft.Office.Interop.Excel.Range;
            //if (currentRange == null) return;


            //MessageBox.Show(currentRange.Font.Color.ToString()); //Color is stored in decimal number 

            //string str = "101";
            //int s = int.Parse(str);
            //MessageBox.Show(s.ToString("X"));

            //MessageBox.Show(int.Parse(currentRange.Font.Color.ToString()).ToString("X"));
            engine = new SpeechRecognitionEngine(new System.Globalization.CultureInfo("en-US"));
            engine.SetInputToDefaultAudioDevice();

            //CultureInfo ci = new CultureInfo("en-us"); //try later
            //sre = new SpeechRecognitionEngine(ci);

            engine.LoadGrammar(new DictationGrammar());//not so correct
            engine.RecognizeAsync(RecognizeMode.Single);

            //engine.SpeechRecognized += Rec;
            //textBox = new System.Windows.Forms.TextBox();
            //textBox.Text = "showing textBox now";
            ////textBox.Visible = true;
            //textBox.Show();
            engine.AudioStateChanged += new EventHandler<AudioStateChangedEventArgs>(AudioChanged);
            engine.SpeechRecognized += new EventHandler<SpeechRecognizedEventArgs>(Rec);

            //currentRange.Value = "st";

            
        }
        private static void AudioChanged(object sender, AudioStateChangedEventArgs e)
        {
            Range currentRange = (Range)Globals.ThisAddIn.Application.Selection as
                Microsoft.Office.Interop.Excel.Range;
            if (currentRange == null) return;
            //currentRange.Value = e.AudioState;  //Silence	1	Receiving silence or non-speech background noise.
            //Speech    2   Receiving speech input.
            //Stopped   0   Not processing audio input.
            //if (e.AudioState != 0)
            //{
            //    currentRange.Value = e.AudioState.ToString();
            //}
            //if ((int)e.AudioState == 1)
            //{
            //    currentRange.Value = e.AudioState.ToString();
            //}
            switch ((int)e.AudioState)
            {
                case 1: currentRange.Value = "Please say something"; break;//NOT SO ACCURATE
                case 2: currentRange.Value = "Listening..."; break;
                //case 0: currentRange.Value = "Stopped"; break;
                default: break;
            }
        }
        private static void Rec(object sender, SpeechRecognizedEventArgs result)
        {
            //Console.WriteLine("you said : {0} conf: {1}", rerult.Result.Text, rerult.Result.Confidence);
            Range currentRange = (Range)Globals.ThisAddIn.Application.Selection as
                Microsoft.Office.Interop.Excel.Range;
            if (currentRange == null) return;
            currentRange.Value = result.Result.Text;
        }


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
