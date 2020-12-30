using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using ExcelDna.Integration;  //class ExcelAsyncUtil

namespace MyExcelAddIn
{
    /// <summary>
    ///     
    /// </summary>
    /// <remarks>Dạng ComAddin. Cài đặt bằng cách chạy Excel/ Mở Excel Addins / Browser và chọn file ExcelDna-Template64.xll hoặc file ExcelDna-Template64-packed.xll </remarks>
    /// <remarks> Bắt buộc phải là public</remarks>
    public class Hust
    {
        /// <summary>
        ///         Hàm Excel EmailSinhVien
        /// </summary>
        /// <param name="HoVaTen">Họ và tên đầy đù bằng tiếng Việt có dấu. Ví dụ Đinh Công Thuật</param>
        /// <param name="MaSoSinhVien">Mã số SV do trường cấp. Ví dụ 20002987</param>
        /// <returns></returns>
        /// <remarks> ExcelDna.Integration.ExcelFunction(Name = ...)  sẽ qui định tên hàm để gọi ra trong Excel </remarks>
        [ExcelDna.Integration.ExcelFunction(Description = "Tính địa chỉ email HUST của sinh viên dựa theo tên và mã số sinh viên", Name = "EmailSinhVien")]
        public static object StudentEmail(
            [ExcelDna.Integration.ExcelArgument(Description ="Họ và tên đầy đủ. Ví dụ Đinh Công Thuật")]  string HoVaTen,  
            [ExcelDna.Integration.ExcelArgument(Description ="Mã số SV do trường cấp. Ví dụ 20002987")]   string MaSoSinhVien)
        {
            string TenKhongDau = NameTools.LayTenKhongDau(HoVaTen);
            string ChuCaiDau = NameTools.LayCacChuCaiDau(TenKhongDau);
            string Email = NameTools.LayTen(TenKhongDau) + "." + ChuCaiDau.Substring(0, ChuCaiDau.Length - 1) + MaSoSinhVien.Substring(2, MaSoSinhVien.Length - 2) + "@sis.hust.edu.vn";
            return Email;
        }

        /// <summary>
        ///         Hàm Excel EmailSinhVien vui tính
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <remarks> Hàm async, cho phép trả về dữ liệu chậm hơn </remarks>
        [ExcelDna.Integration.ExcelFunction(Description = "Tìm địa chỉ email HUST của sinh viên dựa theo tên và mã số sinh viên", Name = "EmailSinhVien2")]
        public static object StudentEmail2(string HoVaTen, string MaSoSinhVien)
        {
            return ExcelAsyncUtil.Run("RunSomethingDelay", new[] { HoVaTen, MaSoSinhVien }, () => RunSomethingDelay(HoVaTen));
        }
        // Hàm async response, 
        public static string RunSomethingDelay(string name)
        {
            Thread.Sleep(1000);
            return $"Hello {name}";
        }

        /// <summary>
        ///         KipThi function, return the starting time of the exam
        /// </summary>
        /// <param name="Kip">Kíp thi từ 1 đến 4</param>
        /// <returns></returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Return starting time of the exam", Name = "KipThi")]
        public static object KipThi(int Kip)
        {
            string startingTime;
            switch (Kip)
            {
                case 1: startingTime = "7:00"; break;
                case 2: startingTime = "9:30"; break;
                case 3: startingTime = "12:30"; break;
                case 4: startingTime = "15:00"; break;
                default: startingTime = "Invalid"; break;
            }
            return startingTime;
        }
        [ExcelDna.Integration.ExcelFunction(Description = "Tên sinh viên làm project 1", Name = "TacGia")]
        static public string TacGia()
        {
            return "Phan Thị Lệ Hằng";
        }
        /// <summary>
        /// Lấy thông tin SoICT
        /// </summary
        /// <returns>6 dòng liên tiếp</returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Thông tin Viện CNTT & TT", Name = "ThongTinSoICT")]
        public static object Laythongtin()
        {
            object[,] result = new object[6, 1];
            result[0, 0] = "VIỆN CÔNG NGHỆ THÔNG TIN VÀ TRUYỀN THÔNG";
            result[1, 0] = "SYMPOSIUM ON INFORMATION AND COMMUNICATION TECHNOLOGY";
            result[2, 0] = "https://soict.hust.edu.vn/";
            result[3, 0] = "https://www.facebook.com/SoictOfficially";
            result[4, 0] = "Nhà B1 – Đại học Bách khoa Hà nội Số 1 Đại Cồ Việt – Hai Bà Trưng – Hà Nội";
            result[5, 0] = "02438692463";

            return result;
        }
        [ExcelDna.Integration.ExcelFunction(Description = "Thông tin của Trường Đại học bách khoa Hà Nội", Name = "ThongTinHust")]
        public static object Laythongtin1()
        {
            object[,] result = new object[6, 1];
            result[0, 0] = "ĐẠI HỌC BÁCH KHOA HÀ NỘI";
            result[1, 0] = "Hanoi University of Science and Technology";
            result[2, 0] = "https://www.hust.edu.vn/";
            result[3, 0] = "https://www.facebook.com/dhbkhanoi";
            result[4, 0] = "Số 1 Đại Cồ Việt, Hai Bà Trưng, Hà Nội";
            result[5, 0] = "024 3869 4242";

            return result;
        }
        ///hàm chuyển đổi khi sign == true
        static public String ConvertDoubleToBin(double d, int bitnum)
        {
            String strD = d.ToString();
            int phanNguyen = (int)d;
            double phandu = d - phanNguyen;
            String beforPoint = Convert.ToString(phanNguyen, 2);
            String afterPoint = "";
            int doDaiChuSoSauDauPhay = bitnum - beforPoint.Length;
            while (doDaiChuSoSauDauPhay != 0)
            {
                phandu *= 2;
                if (phandu > 1)
                {
                    phandu -= 1;
                    afterPoint += "1";
                }
                else if (phandu == 1)
                {
                    afterPoint += "1";
                    break;
                }
                else if (phandu < 1)
                {
                    afterPoint += "0";
                }
                doDaiChuSoSauDauPhay--;
            }
            String temp = "";
            if (doDaiChuSoSauDauPhay != 0)
            {
                for (int i = 0; i < doDaiChuSoSauDauPhay - 1; i++)
                {
                    temp += "0";
                }
            }
            beforPoint = temp + beforPoint;


            return beforPoint + "." + afterPoint;
        }
        [ExcelDna.Integration.ExcelFunction(Description = "chuyển số thập phân sang nhị phân, d: số cần chuyển, bitnum: số bit, sign = true: số có dấu, = false: số không dấu", Name = "Dec2Bin2")]
        static public String Dec2Bin2(double d, int bitnum, bool sign = false)
        {
            if (sign)
            {
                return ConvertDoubleToBin(d, bitnum);
            }
            int decimall = (int)d;
            String result = ""; /// kêt quả convert
            String temp = Convert.ToString(decimall, 2); /// khi convert decimall sang số nhị phân
            int lengTemp = temp.Length; /// số chữ số nhị phân có nghĩa
            int lengResultWithOutTemp = bitnum - lengTemp; // số chữ số nhị phân không có nghĩa, số chữ số 0 ở đầu
            if (lengResultWithOutTemp < 0) // khi độ dài của bit nhị phân có nghĩa lớn hơn độ dài của số bit biểu diễn -> lỗi
            {
                return "#ERROR";
            }
            String s = "";
            for (int i = 0; i < lengResultWithOutTemp; i++) // tạo số các chữ số 0 ở đầu
            {
                s += "0";
            }
            result = s + temp; // cộng kết quả với số các chữ số 0 ở đầu

            return result;
        }
        /// gio hoc
        [ExcelDna.Integration.ExcelFunction(Description = "Nhập vào tiết học, trả về thời gian bắt đầu tiết nếu nhập true, thời gian kết thúc nếu nhập false", Name = "GioHoc")]
        static public String GioHoc(int i, bool a)
        {
            double start1 = 6.75;
            double start2 = 12.5;
            double k = 0.75;
            double t = 0.166666667;
            double tong = 0;
            if (i <= 6)
            {
                if (i == 1 || i == 2) { tong = start1 + k * (i - 1); }
                else if (i == 3) { tong = start1 + k * (i - 1) + t; }
                else if (i == 4) { tong = start1 + k * (i - 1) + 2 * t; }
                else if (i == 5) { tong = start1 + k * (i - 1) + 3 * t; }
                else { tong = start1 + k * 5 + 3 * t; }
            }
            else
            {
                if (i == 7 || i == 8) { tong = start2 + k * (i - 6 - 1); }
                else if (i == 9) { tong = start2 + k * (i - 6 - 1) + t; }
                else if (i == 10) { tong = start2 + k * (i - 6 - 1) + 2 * t; }
                else if (i == 11) { tong = start2 + k * (i - 6 - 1) + 3 * t; }
                else { tong = start2 + k * 5 + 3 * t; }
            }

            if (a == true)
            {
                int h = (int)tong;
                int n = (int)((tong - h) * 60);
                if (n == 0)
                {
                    String hh = h.ToString() + ":00";
                    return hh;
                }
                else
                {
                    String hh = h.ToString() + ":" + n.ToString();
                    return hh;
                }

            }
            else
            {
                tong = tong + 0.75;
                int h = (int)tong;
                int n = (int)((tong - h) * 60);
                if (n == 0)
                {
                    String hh = h.ToString() + ":00";
                    return hh;
                }
                else
                {
                    String hh = h.ToString() + ":" + n.ToString();
                    return hh;
                }
            }
        }
        /// tiet hoc
        /// <summary>
        ///     Nhập thời gian theo kiểu string
        /// </summary>
        /// <param name="text">Thời gian</param>
        /// <returns>Tên</returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Hiển thị tiết học theo thời gian, nhập vào theo định dạng string hh:mm", Name = "Tiethoc")]
        static public int Tiethoc(String x)
        {
            String[] k = x.Split(':');
            int n = Int16.Parse(k[0]);
            int m = Int16.Parse(k[1]);
            if (n >= 6 && n <= 17)
            {
                if ((n == 6 && m >= 45) || (n == 7 && m < 30)) return 1;
                if ((n == 7 && m >= 30) || (n == 8 && m <= 15)) return 2;
                if ((n == 8 && m >= 25) || (n == 9 && m <= 10)) return 3;
                if ((n == 9 && m >= 20) || (n == 10 && m <= 5)) return 4;
                if ((n == 10 && m >= 15) && m < 60) return 5;
                if ((n == 11 && m >= 0 && m <= 45)) return 6;
                if ((n == 12 && m >= 30) || (n == 13 && m < 15)) return 7;
                if ((n == 13 && m >= 15) || (n == 14 && m == 0)) return 8;
                if (n == 14 && m <= 55 && m >= 10) return 9;
                if (n == 15 && m <= 50 && m >= 5) return 10;
                if ((n == 16 && m < 45 && m >= 0)) return 11;
                if ((n == 16 && m >= 45) || (n == 17 && m <= 30)) return 12;
            }
            return 0;
        }
        /// <summary>
        /// quydoidiem10-4
        /// </summary>
        /// <param name="double">điểm</param>
        /// <returns></returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Chuyển đổi thang điểm 10 sang 4, nhập vào thang điểm 10", Name = "Convert10_4")]
        static public double ThangDiem10_4(double x, bool tomau = false)
        {
            if (x >= 8.5 && x <= 10) x = 4.0;
            else if (x >= 8 && x < 8.5) x = 3.5;
            else if (x >= 7 && x < 8) x = 3.0;
            else if (x >= 6.5 && x < 7) x = 2.5;
            else if (x >= 5.5 && x < 6.5) x = 2.0;
            else if (x >= 5 && x < 5.5) x = 1.5;
            else if (x >= 4 && x < 5) x = 1.0;
            else if (x < 4) x = 0.0;
            return x;
        }

    }
}
