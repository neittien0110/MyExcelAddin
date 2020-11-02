using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Xml;

namespace ExcelFuntions
{
    public class FinanceTools
    {
        private static readonly string[] ChuSo = new string[10] { " không", " một", " hai", " ba", " bốn", " năm", " sáu", " bẩy", " tám", " chín" };
        private static readonly string[] Tien = new string[6] { "", " nghìn", " triệu", " tỷ", " nghìn tỷ", " triệu tỷ" };

        private static readonly string[] NumberInText = new string[10] { "zero" , "one", "two", "three", "four", "five", "six", "seven", "eight", "nine"};
        private static readonly string[] Money = new string[6] { "", " thousand", " million", " billion", "trillion", " quadrillion"};

        private static readonly long BiggestNumber = 8999999999999999;

        /// <summary>
        /// Hàm đọc số thành chữ
        /// </summary>
        private static string DocTienBangChu(long SoTien)
        {
            int lan, i;
            long so;
            string KetQua = "", tmp;
            int[] ViTri = new int[6];
            if (SoTien == 0) return "";
            if (SoTien > 0)
            {
                so = SoTien;
            }
            else
            {
                so = -SoTien;
            }
            //Kiểm tra số quá lớn
            if (SoTien > BiggestNumber)
            {
                //SoTien = 0;
                return "";
            }
            ViTri[5] = (int)(so / 1000000000000000);
            so -= long.Parse(ViTri[5].ToString()) * 1000000000000000;
            ViTri[4] = (int)(so / 1000000000000);
            so -= long.Parse(ViTri[4].ToString()) * +1000000000000;
            ViTri[3] = (int)(so / 1000000000);
            so -= long.Parse(ViTri[3].ToString()) * 1000000000;
            ViTri[2] = (int)(so / 1000000);
            ViTri[1] = (int)((so % 1000000) / 1000);
            ViTri[0] = (int)(so % 1000);
            if (ViTri[5] > 0)
            {
                lan = 5;
            }
            else if (ViTri[4] > 0)
            {
                lan = 4;
            }
            else if (ViTri[3] > 0)
            {
                lan = 3;
            }
            else if (ViTri[2] > 0)
            {
                lan = 2;
            }
            else if (ViTri[1] > 0)
            {
                lan = 1;
            }
            else
            {
                lan = 0;
            }
            for (i = lan; i >= 0; i--)
            {
                tmp = DocSo3ChuSo(ViTri[i]);
                KetQua += tmp;
                if (ViTri[i] != 0) KetQua += Tien[i];
                if ((i > 0) && (!string.IsNullOrEmpty(tmp))) KetQua += ",";//&& (!string.IsNullOrEmpty(tmp))
            }
            if (KetQua.Substring(KetQua.Length - 1, 1) == ",") KetQua = KetQua.Substring(0, KetQua.Length - 1);
            KetQua = KetQua.Trim();
            if (SoTien < 0)
            {
                KetQua = "âm " + KetQua;
            }    
            return KetQua.Substring(0, 1).ToUpper() + KetQua.Substring(1);
        }
        ///<summary>
        ///     Function read money in text in English
        /// </summary>
        private static string ReadMoneyInText(long Amount)
        {
            int times, i;
            long number;
            string Result = "", tmp;
            int[] Place = new int[6];
            if (Amount == 0) return "";
            if (Amount > 0)
            {
                number = Amount;
            }
            else
            {
                number = -Amount;
            }
            // Check if the number is too big
            if (Amount > BiggestNumber)
            {
                //Amount = 0;
                return "";
            }
            Place[5] = (int)(number / 1000000000000000);
            number -= long.Parse(Place[5].ToString()) * 1000000000000000;
            Place[4] = (int)(number / 1000000000000);
            number -= long.Parse(Place[4].ToString()) * +1000000000000;
            Place[3] = (int)(number / 1000000000);
            number -= long.Parse(Place[3].ToString()) * 1000000000;
            Place[2] = (int)(number / 1000000);
            Place[1] = (int)((number % 1000000) / 1000);
            Place[0] = (int)(number % 1000);
            if (Place[5] > 0)
            {
                times = 5;
            }
            else if (Place[4] > 0)
            {
                times = 4;
            }
            else if (Place[3] > 0)
            {
                times = 3;
            }
            else if (Place[2] > 0)
            {
                times = 2;
            }
            else if (Place[1] > 0)
            {
                times = 1;
            }
            else
            {
                times = 0;
            }
            for (i = times; i >= 0; i--)
            {
                tmp = Read3DigitNumber(Place[i]);
                Result += tmp;
                if (Place[i] != 0) Result += Money[i];
                if ((i > 0) && (!string.IsNullOrEmpty(tmp))) Result += ", ";//&& (!string.IsNullOrEmpty(tmp))
            }
            if (Result.Substring(Result.Length - 2, 2) == ", ") Result = Result.Substring(0, Result.Length - 2);
            Result = Result.Trim();
            if (Amount < 0)
            {
                Result = "Negative " + Result;
            }
            return Result.Substring(0, 1).ToUpper() + Result.Substring(1);
        }

        /// <summary>
        ///     Hàm đọc số có 3 chữ số
        /// </summary>
        /// <param name="baso"></param>
        /// <returns></returns>
        static private string DocSo3ChuSo(int baso)
        {
            int tram, chuc, donvi;
            string KetQua = "";
            tram = (int)(baso / 100);
            chuc = (int)((baso % 100) / 10);
            donvi = baso % 10;
            if ((tram == 0) && (chuc == 0) && (donvi == 0)) return "";
            if (tram != 0)
            {
                KetQua += ChuSo[tram] + " trăm";
                if ((chuc == 0) && (donvi != 0)) KetQua += " linh";
            }
            if ((chuc != 0) && (chuc != 1))
            {
                KetQua += ChuSo[chuc] + " mươi";
                if ((chuc == 0) && (donvi != 0)) KetQua += " linh";
            }
            if (chuc == 1) KetQua += " mười";
            switch (donvi) 
            {
                case 1:
                    if ((chuc != 0) && (chuc != 1))
                    {
                        KetQua += " mốt";
                    }
                    else
                    {
                        KetQua += ChuSo[donvi];
                    }
                    break;
                case 5:
                    if (chuc == 0)
                    {
                        KetQua += ChuSo[donvi];
                    }
                    else
                    {
                        KetQua += " lăm";
                    }
                    break;
                default:
                    if (donvi != 0)
                    {
                        KetQua += ChuSo[donvi];
                    }
                    break;
            }
            return KetQua;
        }
        /// <summary>
        ///     Function read 3-digit number
        /// </summary>
        /// <param name="threeDigit"></param>
        /// <returns></returns>
        static private string Read3DigitNumber( int threeDigit)
        {
            int hundreds, tens, ones;
            string Result = "";
            hundreds = (int)(threeDigit / 100);
            tens = (int)((threeDigit % 100) / 10);
            ones = threeDigit % 10;
            if ((hundreds == 0) && (tens == 0) && (ones == 0)) return "";
            if (hundreds != 0)
            {
                Result += NumberInText[hundreds] + " hundred";
                //if ((tens == 0) && (ones != 0)) Result += ""
            }
            if ((tens == 0) && (ones != 0))
            {
                Result += " "+ NumberInText[ones];
            }
            else if ((tens == 0) && (ones == 0))
            {

            }
            else
            {
                if (tens == 1)
                {
                    switch (ones)
                    {
                        case 0: Result += " ten"; break;
                        case 1: Result += " eleven"; break;
                        case 2: Result += " twelve"; break;
                        case 3: Result += " thirteen"; break;
                        case 4: Result += " fourteen"; break;
                        case 5: Result += " fifteen"; break;
                        case 6: Result += " sixteen"; break;
                        case 7: Result += " seventeen"; break;
                        case 8: Result += " eighteen"; break;
                        case 9: Result += " nineteen"; break;
                        default: break;
                    }
                }
                else
                {
                    switch (tens)
                    {
                        case 2: Result += " twenty"; break;
                        case 3: Result += " thirty"; break;
                        case 4: Result += " forty"; break;
                        case 5: Result += " fifty"; break;
                        case 6: Result += " sixty"; break;
                        case 7: Result += " seventy"; break;
                        case 8: Result += " eighty"; break;
                        case 9: Result += " ninety"; break;
                        default: break;
                    }
                    Result += "-"+ NumberInText[ones];
                }
            }
            
            return Result;

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Viết số bằng chữ tiếng Việt có dấu")]
        public static string BangChu(
            [ExcelDna.Integration.ExcelArgument(Description = "Số nguyên cần viết bằng chữ")] long number
            )
        {
            return DocTienBangChu(number);
        }
        /// <summary>
        ///     Main function to read number in text
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        public static string InText(
            [ExcelDna.Integration.ExcelArgument(Description = "Integer number to text")] long number
            )
        {
            return ReadMoneyInText(number);
        }



        /// <summary>
        ///     Lấy tỉ giá trừ trang web của Vietcommbank
        /// </summary>
        /// <param name="currency"></param>
        /// <param name="type"></param>
        /// <returns></returns>
        /// <remarks>
        ///     Website được crawl có ngày tháng: https://portal.vietcombank.com.vn/UserControls/TVPortal.TyGia/pListTyGia.aspx?txttungay=20/09/2020&BacrhID=68&isEn=True
        ///     hoặc web được crawl hiện tại:  https://portal.vietcombank.com.vn/Usercontrols/TVPortal.TyGia/pXML.aspx
        /// </remarks>
        [ExcelDna.Integration.ExcelFunction(Description = "Lấy tỉ giá hối đoái ngoại tệ và vnđ theo niêm yết tại portal.vietcombank.com.vn")]
        public static string TyGia(
            [ExcelDna.Integration.ExcelArgument(Description = "Mã ngoại tệ. Giá trị hợp lệ: AUD,CAD,CHF,CNY,DKK,EUR,GBP,HKD,INR,JPY,KRW,KWD,MYR,NOK,RUB,SAR,SEK,SGD,THB,USD")] string currency,
            [ExcelDna.Integration.ExcelArgument(Description = "Loại tỷ giá. Giá trị hợp lệ: mua,bán,chuyển khoản,buy,sell,transfer")] string type,
            [ExcelDna.Integration.ExcelArgument(Description = "Ngày lấy tỷ giá. Cú pháp dd/mm/yyyy. Mặc định là hiện tại", Name = "[Ngày tỷ giá]")] string date = ""
            )
        {
            /// Xác định loại ngoại tệ  
            currency = currency.ToUpper();
            if (currency == "VND")
            {
                return "1";
            }    

            if (!(currency == "AUD" || currency == "CAD" || currency == "CHF" || currency == "CNY" || currency == "DKK" || currency == "EUR" 
                || currency == "GBP" || currency == "HKD" || currency == "INR" || currency == "JPY" || currency == "KRW" || currency == "KWD" 
                || currency == "MYR" || currency == "NOK" || currency == "RUB" || currency == "SAR" || currency == "SEK" || currency == "SGD"
                || currency == "THB" || currency == "USD"  ))
            {
                return "Don't known currency";
            }

            /// Xác định loại ngày lấy tỷ giá
            if (date != "")
            {
#pragma warning disable IDE0059 // Unnecessary assignment of a value
                DateTime ExchangeDate = new DateTime();
#pragma warning restore IDE0059 // Unnecessary assignment of a value
                try
                {
                    ExchangeDate = DateTime.ParseExact(date, "dd/mm/yyyy", null);
                }
                catch
                {
                    return "Invalid date";
                }
                if (ExchangeDate.Date > DateTime.Today)
                {
#pragma warning disable IDE0059 // Unnecessary assignment of a value
                    date = "";
#pragma warning restore IDE0059 // Unnecessary assignment of a value
                    return "Out of date";
                }
                date = ExchangeDate.ToString("dd/mm/yyyy");
            }            

            /// Xác định loại tỷ giá Mua/bán 
            type = type.ToUpper();
            if (type == "MUA" || type == "BUY")
            {
                type = "Buy";
            }
            else if (type == "BÁN" || type == "SELL")
            {
                type = "Sell";
            } else if (type == "CHUYỂN KHOẢN" || type == "TRANSFER")
            {
                type = "Transfer";
            }
            else
            {
                return "Don't known type";
            }

            string reply;
            /// Mở một Webclient
            using (WebClient client = new WebClient())
            {
                ///Khai báo sử dụng SSL
                client.Headers.Add("User-Agent: BrowseAndDownload");
                ServicePointManager.Expect100Continue = true;
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                /// Tạo URL để lấy thông tin từ VCB với tham số ngày tháng
                StringBuilder VCBCrawedURL = new StringBuilder(150);
                if (date == "")
                {
                    VCBCrawedURL.Append("https://portal.vietcombank.com.vn/Usercontrols/TVPortal.TyGia/pXML.aspx");
                } else
                {
                    VCBCrawedURL.AppendFormat("https://portal.vietcombank.com.vn/UserControls/TVPortal.TyGia/pListTyGia.aspx?BacrhID=68&isEn=True&txttungay={0}", date);
                }

                ///Tải về nội dung từ URL của Vietcommbank
                reply = client.DownloadString(VCBCrawedURL.ToString());
                if (date != "")
                {
                    /// URL trả về trang html nên phân tích khá phức tạp
                    HtmlDocument htmlDoc = new HtmlDocument();
                    htmlDoc.LoadHtml(reply);
                    HtmlNodeCollection hExchangedRows;

                    // Html Selector //*[@id=\"ctl00_Content_ExrateView\"]/tbody/tr[12]/td[2]"); //==JPY
                    hExchangedRows = htmlDoc.DocumentNode.SelectNodes("//*[@id=\"ctl00_Content_ExrateView\"]/tbody/tr[@class='odd']"); 
                    foreach (HtmlNode myNode in hExchangedRows)
                    {   // Cấu trúc của myNode
                        //<tr class="odd" data-time="01/03/2020 18:00:00">
                        //    <td style="text-align:left;"> SOUTH KOREAN WON</td>
                        //    <td style="text-align:center;">KRW</td>
                        //    <td>18.41 </td>
                        //    <td>19.38</td>
                        //    <td>20.93</td>
                        //</tr>
                        if (myNode.ChildNodes.Count < 5*2+1) continue;
                        try
                        {
                            if (myNode.ChildNodes[1*2+1].InnerText == currency)
                            {
                                if (type == "Buy")
                                {
                                    reply = myNode.ChildNodes[2 * 2 + 1].InnerText;
                                } else if (type == "Transfer")
                                {
                                    reply = myNode.ChildNodes[3 * 2 + 1].InnerText;
                                } else if (type == "Sell")
                                {
                                    reply = myNode.ChildNodes[4 * 2 + 1].InnerText;
                                }

                            }
                        }
                        catch (Exception ex)
                        {
                            Debug.WriteLine(ex.Message);
                            Debug.WriteLine(myNode.OuterHtml);
                        }
                    }     
                    

                }   else
                {
                    /// URL trả về payload dạng xml nên khá đơn giản
                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(reply);
                    XmlNode xnode;
                    xnode = xmlDoc.DocumentElement.SelectSingleNode("/ExrateList/Exrate[@CurrencyCode='" + currency + "']");
                    reply = xnode.Attributes.GetNamedItem(type).Value;
                }    
                VCBCrawedURL.Clear();
            }
            /// Chuyển đổi về dạng số và có lưu ý về qui tắc dấu .,
#pragma warning disable IDE0017 // Simplify object initialization
            NumberFormatInfo VCBformatProvider = new NumberFormatInfo();
#pragma warning restore IDE0017 // Simplify object initialization
            VCBformatProvider.NumberDecimalSeparator = ".";
            VCBformatProvider.NumberGroupSeparator = ",";
            /// Trả về quả 
            if (reply != "-")
            {
                try
                {
                    return Convert.ToDouble(reply, VCBformatProvider).ToString();
                }
                catch
                {
                    return "No data";
                }
            }
            else
            {   // - Khi VCB không chứa thông tin, họ trả về kí tự -
                return "No data";
            }    
        }
    }
}
