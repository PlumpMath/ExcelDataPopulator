using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataPopulator
{
    class Program
    {
        static void Main(string[] args)
        {
            var sw = Stopwatch.StartNew();

            Examples.RunPopulate();
            //Examples.RunStringArrayConverterPerformance();
            
            Console.WriteLine("Finished in {0} ms", sw.ElapsedMilliseconds);
            sw.Stop();
            Console.Read();
        }




    }


}
