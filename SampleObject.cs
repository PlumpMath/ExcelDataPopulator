using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataPopulator
{
    class SampleObject
    {
        public int RowId { get; set; }
        public string A { get; set; }
        public string B { get; set; }
        public string C { get; set; }
        public string D { get; set; }
        public string E { get; set; }

        public decimal Qty { get; set; }
        public decimal Price { get; set; }
        public decimal Amount { get { return Qty * Price; } }

        public override string ToString()
        {
            return string.Format("{0}, {1}, {2}, {3}, {4}, {5}, {6}, {7}, {8}", RowId, A, B, C, D, E, Qty, Price, Amount);
        }
    }
}
