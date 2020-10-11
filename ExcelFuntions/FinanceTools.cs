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
        static private string[] ChuSo = new string[10] { " không", " một", " hai", " ba", " bốn", " năm", " sáu", " bẩy", " tám", " chín" };
        static private string[] Tien = new string[6] { "", " nghìn", " triệu", " tỷ", " nghìn tỷ", " triệu tỷ" };

        /// <summary>
        /// Hàm đọc số thành chữ
        /// </summary>
        private static string DocTienBangChu(long SoTien)
        {
            int lan, i;
            long so;
            string KetQua = "", tmp = "";
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
            if (SoTien > 8999999999999999)
            {
                SoTien = 0;
                return "";
            }
            ViTri[5] = (int)(so / 1000000000000000);
            so = so - long.Parse(ViTri[5].ToString()) * 1000000000000000;
            ViTri[4] = (int)(so / 1000000000000);
            so = so - long.Parse(ViTri[4].ToString()) * +1000000000000;
            ViTri[3] = (int)(so / 1000000000);
            so = so - long.Parse(ViTri[3].ToString()) * 1000000000;
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
        // Hàm đọc số có 3 chữ số
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
                if ((chuc == 0) && (donvi != 0)) KetQua = KetQua + " linh";
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

        [ExcelDna.Integration.ExcelFunction(Description = "Viết số bằng chữ tiếng Việt có dấu", Category = "Text")]
        public static string BangChu(
            [ExcelDna.Integration.ExcelArgument(Description = "Số nguyên cần viết bằng chữ")] int number
            )
        {
            return DocTienBangChu(number);
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
        [ExcelDna.Integration.ExcelFunction(Description = "Lấy tỉ giá hối đoái ngoại tệ và vnđ theo niêm yết tại portal.vietcombank.com.vn", Category = "Financial")]
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
                DateTime ExchangeDate = new DateTime();
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
                    date = "";
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
            NumberFormatInfo VCBformatProvider = new NumberFormatInfo();
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
