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
        [ExcelDna.Integration.ExcelFunction(Description = "Tính địa chỉ email HUST của sinh viên, giảng viên dựa theo tên và mã số sinh viên", Name = "EmailSinhVien", Category = "Text")]
        public static object StudentEmail(
            [ExcelDna.Integration.ExcelArgument(Description = "Họ và tên đầy đủ. Ví dụ Đinh Công Thuật")] string HoVaTen,
            [ExcelDna.Integration.ExcelArgument(Description = "Mã số SV do trường cấp. Bỏ trống nếu là giảng viên. Ví dụ 20002987.")] string MaSoSinhVien)
        {
            string TenKhongDau = NameTools.LayTenKhongDau(HoVaTen);
            string Email;
            if (MaSoSinhVien.Length == 8)
            {
                string ChuCaiDau = NameTools.LayCacChuCaiDau(TenKhongDau);
                Email = NameTools.LayTen(TenKhongDau) + "." + ChuCaiDau.Substring(0, ChuCaiDau.Length - 1) + MaSoSinhVien.Substring(2, MaSoSinhVien.Length - 2) + "@sis.hust.edu.vn";
            }
            else if (MaSoSinhVien.Length == 0)
            {
                Email = NameTools.LayTen(TenKhongDau) + "." + NameTools.LayHo(TenKhongDau) + NameTools.LayDem(TenKhongDau).Trim() + "@hust.edu.vn";
            }
            else
            {
                Email = "Mã số không hợp lệ";
            }    
            return Email;
        }

        /// <summary>
        ///         Hàm Excel EmailSinhVien vui tính
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        /// <remarks> Hàm async, cho phép trả về dữ liệu chậm hơn </remarks>
        [ExcelDna.Integration.ExcelFunction(Description = "Tìm địa chỉ email HUST của sinh viên dựa theo tên và mã số sinh viên", Name = "EmailSinhVien2", Category = "Text")]
        public static object JustForFun(string HoVaTen, string MaSoSinhVien)
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
        /// Students the profile.
        /// </summary>
        /// <param name="HoVaTen">The ho va ten.</param>
        /// <param name="MaSoSinhVien">The ma so sinh vien.</param>
        /// <returns>
        ///     https://husteduvn-my.sharepoint.com/person.aspx?user=tuong.pd164582@sis.hust.edu.vn
        /// </returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Đường link thông tin cá nhân công khai của email HUST, bao gồm cả giáo viên và sinh viên. ", Category = "Text")]
        public static object SharepointProfile(
         [ExcelDna.Integration.ExcelArgument(Description = "Email HUST của sinh viên. Ví dụ chinh.bq002987@sis.hust.edu.vn")] string EmailSinhVien)
        {
            return $"https://husteduvn-my.sharepoint.com/person.aspx?user={EmailSinhVien}";
        }
        

    }
}
