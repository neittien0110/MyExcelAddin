using System;


namespace MyExcelAddIn
{
    /// <summary>
    ///     Các hàm xử lý chuỗi tên
    /// </summary>
    /// <remarks> Bắt buộc phải là public</remarks>
    public class NameTools
    {
        /// <summary>
        ///     Danh sách các kí tự tiếng Việt
        /// </summary>
        private static readonly string[] VietnameseSigns = new string[]
        {

            "aAeEoOuUiIdDyY",

            "áàạảãâấầậẩẫăắằặẳẵ",

            "ÁÀẠẢÃÂẤẦẬẨẪĂẮẰẶẲẴ",

            "éèẹẻẽêếềệểễ",

            "ÉÈẸẺẼÊẾỀỆỂỄ",

            "óòọỏõôốồộổỗơớờợởỡ",

            "ÓÒỌỎÕÔỐỒỘỔỖƠỚỜỢỞỠ",

            "úùụủũưứừựửữ",

            "ÚÙỤỦŨƯỨỪỰỬỮ",

            "íìịỉĩ",

            "ÍÌỊỈĨ",

            "đ",

            "Đ",

            "ýỳỵỷỹ",

            "ÝỲỴỶỸ"
        };

        /// <summary>
        ///         Trích ra tên không dấu
        /// </summary>
        /// <param name="text">Họ và tên đầy đù bằng tiếng Việt có dấu. Ví dụ Đinh Công Thuật</param>
        /// <returns>Từ cuối cùng</returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Trả về văn bản ở dạng tiếng Việt không dấu", Category = "Text")]
        public static string VanBanKhongDau(
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
        [ExcelDna.Integration.ExcelFunction(Description = "Trả về tên ở dạng tiếng Việt không dấu", Category = "Text")]
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
        [ExcelDna.Integration.ExcelFunction(Description = "Tên sinh viên", Category = "Text")]
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
        [ExcelDna.Integration.ExcelFunction(Description = "Tên sinh viên", Category = "Text")]
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
        [ExcelDna.Integration.ExcelFunction(Description = "Tên sinh viên", Category = "Text")]
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
        [ExcelDna.Integration.ExcelFunction(Description = "Tên sinh viên", Category = "Text")]
        public static string LayCacChuCaiDau(
            [ExcelDna.Integration.ExcelArgument(Description = "Họ và tên đầy đủ. Ví dụ Đinh Công Thuật")] string HoVaTen)
        {
            return NameTools.ExtractFirstLetters(HoVaTen);
        }

        /// <summary>
        ///     Loại bỏ dấu khỏi chuỗi văn bản tiếng Việt
        /// </summary>
        /// <param name="text">Văn bản có dấu </param>
        /// <returns>văn bản không dấu</returns>
        static public string RemoveAccent(string text)
        {
            for (int i = 1; i < VietnameseSigns.Length; i++)
            {
                for (int j = 0; j < VietnameseSigns[i].Length; j++)
                    text = text.Replace(VietnameseSigns[i][j], VietnameseSigns[0][i - 1]);
            }
            return text;
        }

        /// <summary>
        ///     Lấy duy nhất tên khỏi tên đầy đủ
        /// </summary>
        /// <param name="text">Tên đầy đủ</param>
        /// <returns>Tên</returns>
        static public string ExtractLastName(string FullName)
        {
            int i;
            int len;
            /// Loại bỏ kí tự trống ở cuối nếu có
            FullName = FullName.TrimEnd();

            len = FullName.Length;
            /// Tìm vị trí kí tự space cuối cùng
            for (i= len - 1; i>=0; i--)
            {
                if (FullName[i] == ' ')
                    break;
            }
            if (i < 0)
            {
                return String.Empty;
            }
            else
            {
                return FullName.Substring(i+1, len - i-1);
            }
            
        }

        /// <summary>
        ///     Lấy duy nhất tên khỏi tên đầy đủ
        /// </summary>
        /// <param name="text">Tên đầy đủ</param>
        /// <returns>Tên</returns>
        static public string ExtractFirstName(string FullName)
        {
            int i;
            int len;
            /// Loại bỏ kí tự trống ở đầu nếu có
            FullName = FullName.TrimStart();

            len = FullName.Length;
            /// Tìm vị trí kí tự space đầu tiên
            for (i = 1; i < len; i++)
            {
                if (FullName[i] == ' ')
                    break;
            }
            if (i < len)
            {
                return FullName.Substring(0, i);
            }
            else
            {
                return FullName;
            }
        }


        /// <summary>
        ///     Lấy duy nhất đệm  khỏi tên đầy đủ
        /// </summary>
        /// <param name="text">Tên đầy đủ</param>
        /// <returns>Tên</returns>
        static public string ExtractMiddleName(string FullName)
        {
            int pos_end;
            int pos_start;
            int len;
            
            /// Loại bỏ kí tự trống ở cả 2 đầu cuối nếu có
            FullName = FullName.Trim();

            len = FullName.Length;

            /// Tìm vị trí kí tự space đầu tiên
            for (pos_start = 1; pos_start < len; pos_start++)
            {
                if (FullName[pos_start] == ' ')
                    break;
            }
            if (pos_start >= len)
            {
                return String.Empty;
            }

            /// Bỏ qua các kí tự trống
            for (pos_start++; pos_start < len && (FullName[pos_start] == ' '); pos_start++) { };

            /// Tìm vị trí kí tự space cuối cùng
            for (pos_end = len - 1; pos_end > pos_start; pos_end--)
            {
                if (FullName[pos_end] == ' ')
                    break;
            }
            if (pos_end < pos_start)
            {
                return String.Empty;
            }
            else
            {
                return FullName.Substring(pos_start, pos_end - pos_start);
            }
        }

        /// <summary>
        ///     Lấy các chữ cái đầu tên
        /// </summary>
        /// <param name="text">Tên đầy đủ</param>
        /// <returns>Tên</returns>
        static public string ExtractFirstLetters(string FullName)
        {
            int i;
            int len;
            string ShorthenName = String.Empty;

            /// Loại bỏ kí tự trống ở cả 2 đầu cuối nếu có
            FullName = FullName.Trim();

            len = FullName.Length;

            /// Nếu chỉ có 1 kí tự thì xong luôn
            if (len == 1)
            {
                ShorthenName = FullName;
            }
            else
            {
                ShorthenName = ShorthenName + FullName[0];
                for (i = 0; i < len; i++)
                {
                    if ((FullName[i] == ' ') && (i < (len - 1)) && (FullName[i + 1] != ' '))
                    {
                        ShorthenName = ShorthenName + FullName[i + 1];
                        i++;
                    }
                }
            }
            return ShorthenName;
        }
    }
}
