using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Sandbox.OpenXML
{
    public class ExcelHelper
    {
        private const int maxColumns = 16384;
        private string[] columnHeaderNames = new string[maxColumns];
        private string[] names = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

        public ExcelHelper()
        {
            var columnName = string.Empty;

            int firstLetterIndex, secondLetterIndex, thirdLetterIndex, columnHeaderIndex;

            firstLetterIndex = secondLetterIndex = thirdLetterIndex = -1;

            for (columnHeaderIndex = 0; columnHeaderIndex < maxColumns; ++columnHeaderIndex)
            {
                columnName = string.Empty;

                ++firstLetterIndex;

                if (firstLetterIndex == 26)
                {
                    firstLetterIndex = 0;

                    ++secondLetterIndex;

                    if (secondLetterIndex == 26)
                    {
                        secondLetterIndex = 0;

                        ++thirdLetterIndex;
                    }
                }

                if (thirdLetterIndex >= 0) columnName += names[thirdLetterIndex];

                if (secondLetterIndex >= 0) columnName += names[secondLetterIndex];

                if (firstLetterIndex >= 0) columnName += names[firstLetterIndex];

                columnHeaderNames[columnHeaderIndex] = columnName;
            }
        }

        public string GetColumnName(int index)
        {
            return columnHeaderNames[index];
        }
    }
}
