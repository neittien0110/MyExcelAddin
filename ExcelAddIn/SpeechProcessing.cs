using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Speech;
using System.Speech.Synthesis;
using System.Globalization;
using System.Media;
using System.Diagnostics;
using System.Windows.Forms;
using System.IO;
using System.Net;
using NAudio.Wave;
using System.Threading;

namespace ExcelAddIn
{
    [Serializable] 
    public enum CELL_ORDER_TYPE
    {
        INFO,
        WARNING,
        EXCEPTION,
        CRITICAL,
        NONE
    }

    /// <summary>
    ///     Các hàm xử lý về tiếng nói
    /// </summary>
    /// <remarks>
    ///     Là hàm backend, cung cấp tinh năng cho các Ribbon, ContextMenu
    /// </remarks>
    public class SpeechProcessing
    {

        /// <summary>
        ///     Đọc văn bản bằng hàm Cotana. Không cần intenet
        /// </summary>
        /// <param name="text">Văn bản cần đọc</param>
        static public void ReadMeByCotana(string text)
        {
            var synthesizer = new SpeechSynthesizer();
            synthesizer.SetOutputToDefaultAudioDevice();
            var builder = new PromptBuilder();
            builder.StartVoice(new CultureInfo("en-US"));
            builder.AppendText(text);
            builder.EndVoice();
            synthesizer.Speak(builder);
        }

        /// <summary>
        ///     Đọc văn bản tiếng Việt bằng dịch vụ code.responsivevoice.org
        /// </summary>
        /// <param name="text">Văn bản cần đọc</param>
        /// <example>
        ///     https://code.responsivevoice.org/getvoice.php?text=B%E1%BA%A1n%20g%C3%AC%20%C6%A1i&lang=vi&engine=g1&name=&pitch=0.5&rate=0.5&volume=1&key=WGciAW2s&gender=female
        ///     https://code.responsivevoice.org/getvoice.php?text=B%E1%BA%A1n%20g%C3%AC%20%C6%A1i&lang=vi&engine=g1&name=&pitch=0.5&rate=0.5&vol=1&key=HY7lTyiS&gender=female
        /// </example>
        static public void ReadMeByResponsiveVoice(string text)
        {
            /// Tạo url tới dịch vụ
            string URL = @"https://code.responsivevoice.org/getvoice.php?text=" + text + "&lang=vi&engine=g1&name=&pitch=0.5&rate=0.5&volume=1&key=HY7lTyiS&gender=female";

                using (Stream stream = WebRequest.Create(URL).GetResponse().GetResponseStream())
                {
                    byte[] buffer = new byte[32768];
                    int read;

                /// Tạo file tạm
                string TempFile = Path.GetTempFileName() + ".mp3";
                /// Lưu dữ liệu âm thanh vào file tạm    
                FileStream outputFileStream = new FileStream(TempFile, FileMode.Create);
                    while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
                    {
                        outputFileStream.Write(buffer, 0, read);
                    }

                 outputFileStream.Close();

                /// Bật file âm thanh
                IWavePlayer waveOutDevice = new WaveOut();
                AudioFileReader audioFileReader = new AudioFileReader(TempFile);

                waveOutDevice.Init(audioFileReader);
                waveOutDevice.Play();

                /// Phải tự blocking quá trình đọc âm thanh
                while (waveOutDevice.PlaybackState == PlaybackState.Playing)
                {
                    Thread.Sleep(300);
                }
                waveOutDevice.Stop();
                audioFileReader.Dispose();
                waveOutDevice.Dispose();

                /// Xóa file tạm
                if (File.Exists(TempFile))
                {
                    File.Delete(TempFile);
                }
            }


            /*
            WebBrowser webBrowser1 = new WebBrowser();
            webBrowser1.Navigate(URL);

            ?*
            /*
            Process proc = new Process();
            proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
            proc.StartInfo.FileName = "https://code.responsivevoice.org/getvoice.php?text=B%E1%BA%A1n%20g%C3%AC%20%C6%A1i&lang=vi&engine=g1&name=&pitch=0.5&rate=0.5&volume=1&key=WGciAW2s&gender=female";
            proc.StartInfo.RedirectStandardError = false;
            proc.StartInfo.RedirectStandardOutput = false;
            proc.StartInfo.CreateNoWindow = true;
            proc.Start();
            proc.Close();
            */
        }
    }
}
