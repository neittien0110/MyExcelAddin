using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace ExcelFuntions
{
    class StringFormatter
    {
        /// <summary>
        /// Return original text and change cell background color
        /// </summary>
        /// <param name="text"></param>
        /// <param name="CellBackgroundColor"></param>
        /// <returns></returns>
        [ExcelDna.Integration.ExcelFunction(Description = "Change cell background color")]
        public static object StringCellFormatter(
            [ExcelDna.Integration.ExcelArgument(Description = "Original text")] string text,
            [ExcelDna.Integration.ExcelArgument(Description = "cell background color")] string CellBackgroundColor)
        {
            var result = Color.FromName(CellBackgroundColor); 
            if (result.IsKnownColor) //check if the color string is valid
            {
                
            }
            return text;
        }
    }
}
