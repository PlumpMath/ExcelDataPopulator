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
        int[] Parse(int number, int b)
        {
            //var d = number / b;

            //if (d > b)
            //{
            //    results.Add(d - b);
            //    Parse(d - b, b, ref results);
            //}
            //else if (d == 0)
            //{
            //    results.Add(number % b);
            //}
            //else
            //{
            //    results.Add(d - 1);
            //    results.Add(number % b);
            //}

            // do from bottom to top

            var results = new List<int>();

            var m = number % b;

            results.Add(m);

            var remainings = 0;
            while ((remainings = (number / b)) >= b)
            {
                results.Add(0);

                number = number / b;
            }

            if (remainings > 0)
                results.Add(remainings - 1);

            //results.Add(m);

            return results.AsEnumerable().Reverse().ToArray();
        }

        public void Generate()
        {
            //var results = Parse(676, 26);
            //Debug.WriteLine(string.Join("-", results.Select(x => x.ToString()).ToArray()));

            //for (int i = 0; i < 16834; i++)
            //{
            //    var results = Parse(i, 26);
            //    Debug.WriteLine("{0} : {1} : {2}", i, string.Join("-", results.Select(x => x.ToString()).ToArray()), string.Join("-", results.Select(x => (char)(x + 65)).ToArray()));
            //    //Debug.WriteLine("{0} : {1}", i, string.Join("-", results.Select(x => x.ToString()).ToArray()));

            //}

            //foreach (var item in Iterate(0,4))
            //{
            //    if (item.Length == 1 || item.Skip(1).All(x => x == 0))
            //        Debug.WriteLine(string.Join("-", item.Select(x => x.ToString()).ToArray()));
            //}

            //Console.WriteLine();

            //foreach (var item in IterateWithReverse().Take(15))
            //{
            //    Console.Write(string.Join("-", item.Select(x => x.ToString()).ToArray()) + ",");
            //}

            var i = 0;
            foreach (var item in Iterate(0, 5).Take(16384))
            {
                //Console.Write(string.Join("-", item.Select(x => x.ToString()).ToArray()) + ",");
                Debug.WriteLine("{0} : {1} : {2}", i, string.Join("-", item.Select(x => x.ToString()).ToArray()), string.Join("-", item.Select(x => (char)(x + 65)).ToArray()));
                //Debug.WriteLine("{0} : {1}", i, string.Join("-", item.Select(x => x.ToString()).ToArray()));
                i++;
            }

            //var i = 0;
            //foreach (var item in Iterate(65, 26).Take(16384))
            //{
            //    //var columnName = string.Join("", item.Select(x => (char)x).ToArray()).Trim();

            //    //Debug.WriteLine("{0} : {1} : {2}", i, string.Join("-", item.Select(x => x.ToString()).ToArray()), string.Join("-", item.Select(x => (char)(x)).ToArray()));

            //    //if (columnName == "XFD")
            //    //    break;
            //    i++;
            //}

            //foreach (var item in IterateWithReverse(65, 26).Take(10000))
            //{
            //    Debug.WriteLine(string.Join("", item.Select(x => (char)x).ToArray()).Trim());
            //}

            //foreach (var item in IterateWithReverse(1, 11).Take(50))
            //{
            //    Debug.WriteLine(string.Join("-", item.Select(x => x.ToString()).ToArray()));
            //}

            //var sw = Stopwatch.StartNew();


            //Debug.WriteLine(string.Join("-", Iterate(0, 11).Take(10000000).Last().Select(x => x.ToString()).ToArray()));

            //Debug.WriteLine(sw.ElapsedMilliseconds);
            //sw.Restart();

            //Debug.WriteLine(string.Join("-", IterateWithReverse(0, 11).Take(10000000).Last().Select(x => x.ToString()).ToArray()));

            //Debug.WriteLine(sw.ElapsedMilliseconds);
            //sw.Restart();



            //Debug.WriteLine(string.Join("-", Iterate(64, 91).Last().Select(x => (char)x).ToArray()));

            //foreach (var item in Iterate(64,91).Take(1000))
            //{
            //    Debug.WriteLine(string.Join("-", item.Select(x => (char)x).ToArray()).Trim());
            //}

        }

        IEnumerable<int[]> IterateWithReverse(int start = 0, int count = 10, int depth = 10, int[] slots = null)
        {
            if (slots == null) slots = new int[depth];

            for (int i = start; i < start + count; i++)
            {
                slots[depth - 1] = i;
                if (depth > 1)
                    foreach (var x in IterateWithReverse(start, count, depth - 1, slots)) yield return x;
                else
                    yield return slots.Reverse().ToArray();
            }
        }

        IEnumerable<int[]> Iterate(int start = 0, int count = 10, int depth = 1, int[] slots = null)
        {
            if (slots == null) slots = new int[depth];

            for (int i = start; i < start + count; i++)
            {
                //if (slots.Length != 1 && slots.Length == depth && i == 0)
                //{
                //    Debug.WriteLine("Slot: {0}, depth: {1}, {2}", slots.Length, depth, i);
                //    continue;
                //}
                    

                slots[depth - 1] = i;

                if (depth > 1)
                    foreach (var x in Iterate(start, count, depth - 1, slots)) yield return x;
                else
                    yield return slots.Reverse().ToArray();
            }

            if (slots.Length == depth)
                foreach (var x in Iterate(start, count, depth + 1, null)) yield return x;
        }


        public void Generate(ref IList<int> seeds, ref IList<int> results, int max)
        {
            //for (int i3 = 0; i3 < 10; i3++)
            //{
            //    for (int i2 = 0; i2 < 10; i2++)
            //    {
            //        for (int i = 0; i < 10; i++)
            //        {
            //            //Console.Write(i);
            //            var result = i3 * 100 + i2 * 10 + i;
            //            results.Add(result.ToString());

            //            if (result.ToString() == max)
            //                return;
            //        }
            //    }
            //}


            for (int i = 0; i < 10; i++)
            {
                //Console.Write(i);
                var sum = i;
                for (int s = 0; s < seeds.Count; s++)
                {
                    sum += Convert.ToInt32(Math.Pow(10, s));
                }

                results.Add(sum);

                if (sum >= max)
                    return;
            }

            seeds.Add(seeds.Count + 1);

            Generate(ref seeds, ref results, max);
        }
    }
}
