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
            ///Cho phép hiển thị các dòng gợi ý hàm và gợi ý tham số của các thuộc tính
            ///<see cref="ExcelDna.Integration.ExcelFunctionAttribute"/> và <see cref="ExcelDna.Integration.ExcelArgumentAttribute"/>
            IntelliSenseServer.Install();

            // There are various options for wrapping and transforming your functions
            // See the Source\Samples\Registration.Sample project for a comprehensive example
            // Here we just change the attribute before registering the functions
            ExcelRegistration.GetExcelFunctions()
                             .Select(UpdateHelpTopic)
                             .RegisterFunctions();
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

        /// <summary>
        ///      Thiết lập url mặc định trợ giúp các hàm
        /// </summary>
        /// <param name="funcReg"></param>
        /// <returns></returns>
        public ExcelFunctionRegistration UpdateHelpTopic(ExcelFunctionRegistration funcReg)
        {
            funcReg.FunctionAttribute.HelpTopic = "http://techlinkvn.com";
            return funcReg;
        }
    }
}
