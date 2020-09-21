using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFuntions
{
    public class ExcelDNAExamples
    {
        [ExcelFunction(IsMacroType = true)]
        public static object GetTheXllName()
        {
            return XlCall.Excel(XlCall.xlfGetWorkspace, 44);
        }


        /// <summary>
        ///      Sắp xếp tăng dần mảng đầu vào
        /// </summary>
        /// <param name="vector"></param>
        /// <returns>Bất kể vector đầu vào là dòng hay cột, kết quả luôn là 1 dòng ngang</returns>
        [ExcelFunction(Name = "DNA.SortVector")]
        public static double[] SortVector(double[] vector)
        {
            Array.Sort(vector);
            return vector;
        }

        /// <summary>
        ///     Hàm trả về một mảng (array fomular), sẽ lưu trong 6 cell ở 2 dòng 3 cột
        /// </summary>
        /// <returns></returns>
        [ExcelFunction(Name = "DNA.GetArray")]
        public static object GetArray()
        {
            return new object[,] { { 3,2,5 }, { "Bùi","Đức", "Bình"} };
        }


        [ExcelFunction(Description = "Hàm luôn ghi nhớ giá trị lớn nhất đã từng nhập vào một cell chỉ định." +
            "Kể cả nếu đã đóng file excel và mở lại vẫn lưu trữ được.",
            IsMacroType = true, Name = "DNA.IncreaseValue"
            )]
        public static double IncreaseValue(
            [ExcelArgument(Description = "Nên là tham chiếu tới 1 cell khác. Nếu là hằng số thì vô tác dụng.")]
            double newValue)
        {
            ExcelReference reference = (ExcelReference)XlCall.Excel(XlCall.xlfCaller);
            object val = reference.GetValue();
            if (val is double && (double)val > newValue)
                return (double)val;
            return newValue;
        }


        // This is a macro that sets a block of cells
        [ExcelFunction(
            Description = "This is a macro that sets a block of cells",
            Name ="DNA.SetSome",
            Category = "AutoSum",  // là các mục trong Ribbon /  Formulas / Function Library group. 
            HelpTopic ="http://soict.hust.edu.vn", // hiển khi khi dùng Function Dialog để trợ giúp viết hàm
            ExplicitRegistration = true
            )]
        public static void SetSome()
        {
            ExcelReference r = new ExcelReference(2, 5, 3, 6);
            bool ok = r.SetValue(new object[,] { { 3.4, 8.9 }, { "Wow!", ExcelError.ExcelErrorValue } });
        }
    }
}
