using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataPopulator
{
    class NumberGenerator
    {
        public void Generate()
        {
            var i = 0;

            foreach (var item in Iterate(0, 26))
            {
                var columnName = string.Join("-", item.Select(x => (char)(x + 65)).ToArray()).Trim();

                    Debug.WriteLine("#{0} = {1}", i, columnName);
                i++;
            }

            //foreach (var item in Iterate(65, 26).Take(16384))
            //{
            //    var columnName = string.Join("-", item.Select(x => (char)x).ToArray()).Trim();
            //    //Debug.WriteLine(columnName);
            //}

            //i = 0;
            //foreach (var item in Iterate(0, 10).Take(16384))
            //{
            //    var columnName = string.Join("-", item.Select(x => (char)(x + 65)).ToArray()).Trim();
            //    Debug.WriteLine("{0} : {1} : {2}", i, string.Join("-", item.Select(x => x.ToString()).ToArray()), columnName);
            //    i++;
            //}

            //i = 0;
            //foreach (var item in Iterate(0, 26).Take(16384))
            //{
            //    var columnName = string.Join("-", item.Select(x => (char)(x + 65)).ToArray()).Trim();
            //    Debug.WriteLine("{0} : {1} : {2}", i, string.Join("-", item.Select(x => x.ToString()).ToArray()), columnName);
            //    i++;
            //}

            //i = 0;
            //foreach (var item in Iterate(65, 26).Take(16384))
            //{
            //    var columnName = string.Join("", item.Select(x => (char)x).ToArray()).Trim();
            //    Debug.WriteLine("{0} : {1} : {2}", i, string.Join("-", item.Select(x => x.ToString()).ToArray()), columnName);
            //    i++;
            //}
        }

        /// <summary>
        /// This is an infinite list of excel kind of sequence, the caller will be controlling how much of data that it really requires
        /// The method signature is very similar to Enumerable.Range(start, count)
        /// 
        /// Whereas, the other 2 optionals parameters are not for the caller, they are only for internal recursive call
        /// </summary>
        /// <param name="start">The value of the first integer in the sequence.</param>
        /// <param name="count">The number of sequential integers to generate.</param>
        /// <param name="depth">NOT TO BE SPECIFIED DIRECTLY, used for recursive purpose</param>
        /// <param name="slots">NOT TO BE SPECIFIED DIRECTLY, used for recursive purpose</param>
        /// <returns></returns>
        IEnumerable<int[]> Iterate(int start, int count, int depth = 1, int[] slots = null)
        {
            if (slots == null) slots = new int[depth];

            for (int i = start; i < start + count; i++)
            {
                slots[depth - 1] = i;

                if (depth > 1)
                    foreach (var x in Iterate(start, count, depth - 1, slots)) yield return x;
                else
                    yield return slots.Reverse().ToArray();
            }

            if (slots.Length == depth)
                foreach (var x in Iterate(start, count, depth + 1, null)) yield return x;
        }
    }
}
