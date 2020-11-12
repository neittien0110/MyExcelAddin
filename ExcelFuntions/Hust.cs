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
    }
}
