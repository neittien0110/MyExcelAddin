using System.Diagnostics;
using System.Linq;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Registration;

namespace ExcelDna.XFunctions
{
    /// <summary>
    ///     Lớp khởi tạo toàn cục, tự động kích hoạt khi Addin được nạp, cho phép tác động lên cả các hàm Excel đã viết
    /// </summary>
    public class AddIn : IExcelAddIn
    {

        public void AutoOpen()
        {
            var functions = ExcelRegistration.GetExcelFunctions().ToList();
            if (HasNativeXMatch())
            {
                foreach (var func in functions)
                {
                    /// Thêm prefix cho tên của tất cả các hàm, giúp dễ tìm
                    func.FunctionAttribute.Name = "HUST." + func.FunctionAttribute.Name ;
                    /// Link hướng dẫn sử dụng hàm
                    func.FunctionAttribute.HelpTopic = "https://users.soict.hust.edu.vn/sinno/projects/myexceladdin/";
                }
            }
            functions.RegisterFunctions();

            ///Cho phép hiển thị các dòng gợi ý hàm và gợi ý tham số của các thuộc tính
            ///<see cref="ExcelDna.Integration.ExcelFunctionAttribute"/> và <see cref="ExcelDna.Integration.ExcelArgumentAttribute"/>
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            //Gỡ bỏ
            IntelliSenseServer.Uninstall();
        }

        bool HasNativeXMatch()
        {
            int xlfXMatch = 620;
            var retval = XlCall.TryExcel(xlfXMatch, out var _, 1, 1);
            return (retval == XlCall.XlReturn.XlReturnSuccess);
        }
    }
}
