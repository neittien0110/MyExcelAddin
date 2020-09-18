using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using ExcelDna.Integration;  //class ExcelAsyncUtil
using MyUtilities;

namespace MyExcelAddIn
{
    /// <summary>
    ///     
    /// </summary>
    /// <remarks>Dạng ComAddin. Cài đặt bằng cách chạy Excel/ Mở Excel Addins / Browser và chọn file ExcelDna-Template64.xll hoặc file ExcelDna-Template64-packed.xll </remarks>
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
            string TenKhongDau = LayTenKhongDau(HoVaTen);
            string ChuCaiDau = LayCacChuCaiDau(TenKhongDau);
            string Email = LayTen(TenKhongDau) + "." + ChuCaiDau.Substring(0, ChuCaiDau.Length - 1) + MaSoSinhVien.Substring(2, MaSoSinhVien.Length - 2) + "@sis.hust.edu.vn";
            return Email;
        }

        /// <summary>
        ///         Trích ra tên không dấu
        /// </summary>
        /// <param name="text">Họ và tên đầy đù bằng tiếng Việt có dấu. Ví dụ Đinh Công Thuật</param>
        /// <returns>Từ cuối cùng</returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Trả về văn bản ở dạng tiếng Việt không dấu")]
        public static string VanBangKhongDau(
            [ExcelDna.Integration.ExcelArgument(Description = "Văn bản tiếng Việt có dấu. Ví dụ Lê Văn Long, ")] string text
            )
        {
            return NameTools.RemoveAccent(text);
        }

        /// <summary>
        ///         Trích ra tên không dấu
        /// </summary>
        /// <param name="HoTen">Họ và tên đầy đù bằng tiếng Việt có dấu. Ví dụ Đinh Công Thuật</param>
        /// <returns>Từ cuối cùng</returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Trả về tên ở dạng tiếng Việt không dấu")]
        public static string LayTenKhongDau(
            [ExcelDna.Integration.ExcelArgument(Description = "Tên có có dấu. Ví dụ Lê Văn Long, ")] string HoTen
            )
        {
            return NameTools.RemoveAccent(HoTen);
        }

        /// <summary>
        ///         Trích ra tên sinh viên từ họ và tên đầy đủ
        /// </summary>
        /// <param name="HoVaTen">Họ và tên đầy đù bằng tiếng Việt có dấu. Ví dụ Đinh Công Thuật</param>
        /// <returns>Ví dụ Thuật</returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Tên sinh viên")]
        public static string LayTen(
            [ExcelDna.Integration.ExcelArgument(Description = "Họ và tên đầy đủ. Ví dụ Đinh Công Thuật")] string HoVaTen)
        {
            return NameTools.ExtractLastName(HoVaTen);
        }

        /// <summary>
        ///         Trích ra họ sinh viên từ họ và tên đầy đủ
        /// </summary>
        /// <param name="HoVaTen">Họ và tên đầy đù bằng tiếng Việt có dấu. Ví dụ Đinh Công Thuật</param>
        /// <returns>Ví dụ Đinh </returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Tên sinh viên")]
        public static string LayHo(
            [ExcelDna.Integration.ExcelArgument(Description = "Họ và tên đầy đủ. Ví dụ Đinh Công Thuật")] string HoVaTen)
        {
            return NameTools.ExtractFirstName(HoVaTen);
        }

        /// <summary>
        ///         Trích ra họ sinh viên từ họ và tên đầy đủ
        /// </summary>
        /// <param name="HoVaTen">Họ và tên đầy đù bằng tiếng Việt có dấu. Ví dụ Đinh Công Thuật</param>
        /// <returns>Ví dụ Công</returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Tên sinh viên")]
        public static string LayDem(
            [ExcelDna.Integration.ExcelArgument(Description = "Họ và tên đầy đủ. Ví dụ Đinh Công Thuật")] string HoVaTen)
        {
            return NameTools.ExtractMiddleName(HoVaTen);
        }

        /// <summary>
        ///         Trích ra các chữ cái đầu tiên của tên
        /// </summary>
        /// <param name="HoVaTen">Họ và tên đầy đù bằng tiếng Việt có dấu. Ví dụ Đinh Công Thuật</param>
        /// <returns>Ví dụ ĐCT</returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Tên sinh viên")]
        public static string LayCacChuCaiDau(
            [ExcelDna.Integration.ExcelArgument(Description = "Họ và tên đầy đủ. Ví dụ Đinh Công Thuật")] string HoVaTen)
        {
            return NameTools.ExtractFirstLetters(HoVaTen);
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
    }
}
