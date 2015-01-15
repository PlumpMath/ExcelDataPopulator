using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataPopulator
{
    class ExcelColumnNames
    {
        static string[] _indexes;

        static ExcelColumnNames()
        {
            _indexes = NumberGenerator
                        .Iterate(65, 26)
                        .Take(16384)
                        .Select(item => string.Join("", item.Select(x => (char)x).ToArray()).Trim())
                        .ToArray();
        }

        public static string GetColumnName(int index)
        {
            return _indexes[index];
        }
    }
}
