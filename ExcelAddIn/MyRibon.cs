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

        /// <summary>
        ///     Trả về nội dung cell đang được chọn
        /// </summary>
        /// <returns></returns>
        static private dynamic GetSellectedCell()
        {
            return Globals.ThisAddIn.Application.Selection();
        }

        /// <summary>
        ///     Trả về các cell đang được chọn
        /// </summary>
        /// <returns></returns>
        static private dynamic GetSellectedCells()
        {
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;

            if (wb == null)
            {
                MessageBox.Show("Bạn phải mở một workbook.");
                return null;
            }

            /// Lấy sheet đang được active hiện thời
            Worksheet ws = wb.ActiveSheet;

            // Lẩy ra nội dung của cell hiện thời đang được chọn
            Range res = (Range)Globals.ThisAddIn.Application.Selection as Microsoft.Office.Interop.Excel.Range;
            return res;
        }

        private void buttonImage2Cells_Click(object sender, RibbonControlEventArgs e)
        {   //load image into cells - by MSc Tien
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
                if (img.Width * img.Height > MAX_PIXEL)
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

            // Lẩy ra nội dung của cell hiện thời đang được chọn
            //((dynamic)ws.Application.ActiveCell).Value.ToString()

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

        private void buttonColorize_Click(object sender, RibbonControlEventArgs e)
        {   //Colorize the cells based on selected color and saturation - by Stnd Tuong

            String color = dropDownColorRGB.SelectedItem.OfficeImageId;
            int saturationInt;
            if (!int.TryParse(editSaturationPeak.Text, out saturationInt)) //boolean to check if saturation is a valid number
            {
                return;
            }

            bool inverted = checkBoxInvert.Checked;
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

            byte realcolor; // the background color of cell

            foreach (Range mycell in currentRange)
            {
                realcolor = Convert.ToByte(255 * mycell.Value / saturationInt);
                if (!inverted)
                {
                    switch (color) // color to display with saturation 
                    {
                        case "AppointmentColor1": mycell.Interior.Color = Color.FromArgb(realcolor, 0, 0); break; //red
                        case "AppointmentColor2": mycell.Interior.Color = Color.FromArgb(0, 0, realcolor); break; //green
                        case "AppointmentColor3": mycell.Interior.Color = Color.FromArgb(0, realcolor, 0); break; //blue
                        case "AppointmentColor4": mycell.Interior.Color = Color.FromArgb(realcolor, realcolor, realcolor); break; //gray
                        default: break;
                    }
                    if (realcolor > 180) // change font color to black or white based on background's saturation
                    {
                        mycell.Font.Color = Color.Black;
                    }
                    else
                    {
                        mycell.Font.Color = Color.White;
                    }
                }
                else  // inverted color
                {
                    byte invlevel = (byte)(255 - realcolor);
                    switch (color) // color to display with saturation 
                    {
                        case "AppointmentColor1": mycell.Interior.Color = Color.FromArgb(255, invlevel, invlevel); break;//red
                        case "AppointmentColor2": mycell.Interior.Color = Color.FromArgb(invlevel, invlevel, 255); break; //green
                        case "AppointmentColor3": mycell.Interior.Color = Color.FromArgb(invlevel, 255, invlevel); break; //blue
                        case "AppointmentColor4": mycell.Interior.Color = Color.FromArgb(invlevel, invlevel, invlevel); break; //gray
                        default: break;
                    }
                } 
            }           
            //test code:
            //currentRange.Value = color;
            //currentRange.Interior.Color = Color.FromArgb(0, 200, 0);
        }

        private static SpeechRecognitionEngine engine;
        //private System.Windows.Forms.TextBox textBox;
        private void buttonCortana_Click(object sender, RibbonControlEventArgs e)
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

        /// <summary>
        ///     Đọc tiếng Anh bằng hàm tổng hợp của Cotana. Không cần internet
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="e">The <see cref="RibbonControlEventArgs"/> instance containing the event data.</param>
        private void buttonTextToSpeech_Click(object sender, RibbonControlEventArgs e)
        {
            /// Lấy ra vùng cell đang được select
            Range CurrentRange = GetSellectedCells();
            if (CurrentRange == null) return;

            /// Lần lượt đọc từng cell
            if (Properties.Settings.Default.CELL_ORDER == "select")
            {
                /// Lần lượt đọc từng cell theo thứ tự select
                foreach (var mycell in CurrentRange.Cells)
                {
                    if (mycell != null && ((dynamic)(mycell)).Value != null)  // cần loại bỏ trường hợp cell rỗng
                    {
                        SpeechProcessing.ReadMeByCotana(((dynamic)(mycell)).Value.ToString());
                    }
                }
            }
            else if (Properties.Settings.Default.CELL_ORDER == "row")
            {

                /// Lần lượt đọc từng cell
                for (int row = 1; row <= CurrentRange.Rows.Count; row++)
                    for (int col = 1; col <= CurrentRange.Columns.Count; col++)
                    {
                        var mycell = CurrentRange.Cells[row, col];
                        if (mycell != null && ((dynamic)(mycell)).Value != null)  // cần loại bỏ trường hợp cell rỗng
                        {
                            SpeechProcessing.ReadMeByCotana(((dynamic)(mycell)).Value.ToString());
                        }
                    }
            }
            else if (Properties.Settings.Default.CELL_ORDER == "column")
            {

                /// Lần lượt đọc từng cell ưu tiên cột trước
                for (int col = 1; col <= CurrentRange.Columns.Count; col++)
                    for (int row = 1; row <= CurrentRange.Columns.Count; row++)
                    {
                        var mycell = CurrentRange.Cells[row, col];
                        if (mycell != null && ((dynamic)(mycell)).Value != null)  // cần loại bỏ trường hợp cell rỗng
                        {
                            SpeechProcessing.ReadMeByCotana(((dynamic)(mycell)).Value.ToString());
                        }
                    }
            }
        }

        private void buttonSpeechVN_Click(object sender, RibbonControlEventArgs e)
        {
            /// Lấy ra vùng cell đang được select
            Range CurrentRange = GetSellectedCells();
            if (CurrentRange == null) return;

            /*
            if (Properties.Settings.Default.CELL_ORDER == "row")
            {
                CurrentRange.Sort(XlSortOrientation.xlSortRows);
                CurrentRange.ReadingOrder()
            }
            else if (Properties.Settings.Default.CELL_ORDER == "column")
            {
                CurrentRange.Sort(XlSortOrientation.xlSortColumns);
            }
            */

            Debug.WriteLine("Selected Range has {0} columns, {1} rows", CurrentRange.Columns.Count, CurrentRange.Rows.Count);

            if (Properties.Settings.Default.CELL_ORDER == "select")
            {
                /// Lần lượt đọc từng cell theo thứ tự select
                foreach (var mycell in CurrentRange.Cells)
                {
                    if (mycell != null && ((dynamic)(mycell)).Value != null)  // cần loại bỏ trường hợp cell rỗng
                    {
                        SpeechProcessing.ReadMeByResponsiveVoice(((dynamic)(mycell)).Value.ToString());
                    }
                }
            }
            else if (Properties.Settings.Default.CELL_ORDER == "row")
            {

                /// Lần lượt đọc từng cell
                for (int row = 1; row <= CurrentRange.Rows.Count; row++)
                    for (int col = 1; col <= CurrentRange.Columns.Count; col++)
                    {
                        var mycell = CurrentRange.Cells[row, col];
                        if (mycell != null && ((dynamic)(mycell)).Value != null)  // cần loại bỏ trường hợp cell rỗng
                        {
                            SpeechProcessing.ReadMeByResponsiveVoice(((dynamic)(mycell)).Value.ToString());
                        }
                    }
            }
            else if (Properties.Settings.Default.CELL_ORDER == "column")
            {

                /// Lần lượt đọc từng cell ưu tiên cột trước
                for (int col = 1; col <= CurrentRange.Columns.Count; col++)
                    for (int row = 1; row <= CurrentRange.Columns.Count; row++)
                    {
                        var mycell = CurrentRange.Cells[row, col];
                        if (mycell != null && ((dynamic)(mycell)).Value != null)  // cần loại bỏ trường hợp cell rỗng
                        {
                            SpeechProcessing.ReadMeByResponsiveVoice(((dynamic)(mycell)).Value.ToString());
                        }
                    }
            }

        }

        void groupAudio_DialogLauncherClick(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
        {
            AdvancedAudioSettingsDialog dlg = new AdvancedAudioSettingsDialog();
            dlg.ShowDialog();
        }


    }
}
