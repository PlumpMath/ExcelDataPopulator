using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataPopulator
{
    class SampleObject
    {
        public string A { get; set; }
        public string B { get; set; }
        public string C { get; set; }
        public string D { get; set; }
        public string E { get; set; }

        public override string ToString()
        {
            return string.Format("{0}, {1}, {2}, {3}, {4}", A, B, C, D, E);
        }
    }
}
