using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace MyUtilities
{
    /// <summary>
    ///     Các hàm xử lý chuỗi tên
    /// </summary>
    class NameTools
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
