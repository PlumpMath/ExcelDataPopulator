using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataPopulator
{
    class LogHelper
    {
        public static void LogFormat(string format, params string[] message)
        {
            Console.WriteLine(string.Format("{0} - {1}", DateTime.Now.ToString("u"), string.Format(format, message)));
        }
    }
}
