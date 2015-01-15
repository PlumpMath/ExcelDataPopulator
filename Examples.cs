using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Linq.Expressions;

namespace ExcelDataPopulator
{
    class Examples
    {
        public static void RunStringArrayConverterPerformance()
        {
            var sw = Stopwatch.StartNew();

            var items = Enumerable.Range(1, 1000000).Select(i => new SampleObject { A = "A" + i, B = "B" + i, C = "C" + i, D = "D" + i, E = "E" + i }).ToList();
            string[,] stringArray = null;
            object[,] objectArray = null;

            LogHelper.LogFormat("Initialized {0} rows, {1} ms", items.Count.ToString(), sw.ElapsedMilliseconds.ToString());
            sw.Restart();

            for (int i = 0; i < 10; i++)
                stringArray = ArrayConverter.ConvertTo2DStringArrayStatically(items);

            LogHelper.LogFormat("Convert with string array statically is completed in {0} ms", sw.ElapsedMilliseconds.ToString());
            sw.Restart();

            //for (int i = 0; i < 10; i++)
            //    stringArray = ArrayConverter.ConvertTo2DStringArrayWithReflection(items);

            //LogHelper.LogFormat("Convert with string array reflection is completed in {0} ms", sw.ElapsedMilliseconds.ToString());
            //sw.Restart();

            //for (int i = 0; i < 10; i++)
            //    stringArray = ArrayConverter.ConvertTo2DStringArrayWithExpressionTree<SampleObject>()(items);

            //LogHelper.LogFormat("Convert with string array expression tree is completed in {0} ms", sw.ElapsedMilliseconds.ToString());
            //sw.Restart();

            for (int i = 0; i < 10; i++)
                objectArray = ArrayConverter.ConvertTo2DObjectArrayStatically(items);

            LogHelper.LogFormat("Convert with statically is completed in {0} ms", sw.ElapsedMilliseconds.ToString());
            sw.Restart();

            //for (int i = 0; i < 10; i++)
            //    objectArray = ArrayConverter.ConvertTo2DObjectArrayWithReflection(items);

            //LogHelper.LogFormat("Convert with reflection is completed in {0} ms", sw.ElapsedMilliseconds.ToString());
            //sw.Restart();

            //for (int i = 0; i < 10; i++)
            //    objectArray = ArrayConverter.ConvertTo2DObjectArrayWithExpressionTree<SampleObject>()(items);

            //LogHelper.LogFormat("Convert with expression tree is completed in {0} ms", sw.ElapsedMilliseconds.ToString());
            //sw.Restart();

            //new NumberGenerator().Generate();
            //}
            //Console.WriteLine(stringArray[items.Count - 1, 0]);

            //RunPopulate();

            //var obj = new SampleObject { A = "StringA", B = "StringB", C = "StringC", D = "StringD", E = "StringE" };
            //PrintObjectStatically(obj);
            //PrintObjectWithReflection(obj);
            //PrintObjectWithExpressionTree(obj);

            Console.WriteLine("Finished in {0} ms", sw.ElapsedMilliseconds);
            sw.Stop();
            Console.Read();
        }

        public static void RunPopulate()
        {
            var rnd = new Random();

            // prepare data source
            var items = Enumerable.Range(1, 1000000)
                                  .Select(i => new SampleObject
                                  {
                                      RowId = i,
                                      A = "A" + rnd.Next(1, 5),
                                      B = "B" + rnd.Next(1, 5),
                                      C = "C" + rnd.Next(1, 5),
                                      D = "D" + rnd.Next(1, 5),
                                      E = "E" + rnd.Next(1, 5),
                                      Qty = Convert.ToDecimal(rnd.NextDouble() * 100),
                                      Price = Convert.ToDecimal(rnd.NextDouble() * 1000)
                                  }).ToList();

            string filePath = null;

            // uncomment 1 option below to verify the performance and results
            //ExcelAPIHelper.CreateNew(WaysToAssignValues);
            //ExcelAPIHelper.CreateNew(ws => PopulateItemsWithSingleCellMethod(ws, items));
            //ExcelAPIHelper.CreateNew(ws => PopulateItemsWithMultiCellsOneRowMethod(ws, items));
            //filePath = ExcelAPIHelper.CreateNew(ws => PopulateItemsWithMultiCellsAndRowsMethod(ws, items));
            //filePath = ExcelAPIHelper.CreateNew(ws => PopulateItemsWithMultiCellsAndRowsMethodWithReflection(ws, items));
            filePath = ExcelAPIHelper.CreateNew(2, (b, sheets) => PopulateItemsWithMultiCellsAndRowsMethodWithExpressionTree(b, sheets, items));

            if (!string.IsNullOrWhiteSpace(filePath))
                Process.Start(filePath);
            Console.Read();

            // clean up
            //File.Delete(filePath);
        }

        static void WaysToAssignValues(Worksheet worksheet)
        {
            LogHelper.LogFormat("Ways to populate values into excel");
            // DO NOT do this as it will create memory leak (excel stays as background process, visible in task manager)
            // best pratice is to hold the object into a variable then call dispose (release com object)
            //ws.get_Range("C5").Value = "ABC";

            var r1 = worksheet.get_Range("A1"); // single cell
            r1.Value = "A1";
            ExcelAPIHelper.DisposeCOMObject(r1);

            var r2 = worksheet.get_Range("A2:C2"); // multi cells in 1 row
            r2.Value = new[] { "A2", "B2", "C2" };
            ExcelAPIHelper.DisposeCOMObject(r2);

            var r3 = worksheet.get_Range("A3:C4"); // multi cells/rows
            r3.Value = new string[,] { { "A3", "B3", "C3" }, { "A4", "B4", "C4" } }; ;
            ExcelAPIHelper.DisposeCOMObject(r3);
        }

        static void PopulateItemsWithSingleCellMethod(Worksheet worksheet, IList<SampleObject> items)
        {
            LogHelper.LogFormat("Start populating {0} rows using single cell method", items.Count.ToString());
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var rowIndex = i + 1;
                var r1 = worksheet.get_Range("A" + rowIndex); r1.Value = item.A; ExcelAPIHelper.DisposeCOMObject(r1);
                var r2 = worksheet.get_Range("B" + rowIndex); r2.Value = item.B; ExcelAPIHelper.DisposeCOMObject(r2);
                var r3 = worksheet.get_Range("C" + rowIndex); r3.Value = item.C; ExcelAPIHelper.DisposeCOMObject(r3);
                var r4 = worksheet.get_Range("D" + rowIndex); r4.Value = item.D; ExcelAPIHelper.DisposeCOMObject(r4);
                var r5 = worksheet.get_Range("E" + rowIndex); r5.Value = item.E; ExcelAPIHelper.DisposeCOMObject(r5);
            }
        }

        static void PopulateItemsWithMultiCellsOneRowMethod(Worksheet worksheet, IList<SampleObject> items)
        {
            LogHelper.LogFormat("Start populating {0} rows using multi cells (per 1 row) method", items.Count.ToString());
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var rowIndex = i + 1;
                var r = worksheet.get_Range(string.Format("A{0}:E{0}", rowIndex));
                r.Value = new[] { item.A, item.B, item.C, item.D, item.E };
                ExcelAPIHelper.DisposeCOMObject(r);
            }
        }

        static void PopulateItemsWithMultiCellsAndRowsMethod(Worksheet worksheet, IList<SampleObject> items)
        {
            LogHelper.LogFormat("Start populating {0} rows using multi cells and rows method", items.Count.ToString());

            var stringArray = ArrayConverter.ConvertTo2DObjectArrayStatically(items);

            LogHelper.LogFormat("Completed converting the data into string array");

            var r = worksheet.get_Range(string.Format("A{0}:E{1}", 1, items.Count + 1));
            r.Value = stringArray;
            ExcelAPIHelper.DisposeCOMObject(r);
        }

        static void PopulateItemsWithMultiCellsAndRowsMethodWithReflection(Worksheet worksheet, IList<SampleObject> items)
        {
            LogHelper.LogFormat("Start populating {0} rows using multi cells and rows method with reflection", items.Count.ToString());

            var stringArray = ArrayConverter.ConvertTo2DObjectArrayWithReflection(items);

            LogHelper.LogFormat("Completed converting the data into string array");

            var r = worksheet.get_Range(string.Format("A{0}:E{1}", 1, items.Count + 1));
            r.Value = stringArray;
            ExcelAPIHelper.DisposeCOMObject(r);
        }

        static void PopulateItemsWithMultiCellsAndRowsMethodWithExpressionTree(Workbook book, IList<Worksheet> worksheets, IList<SampleObject> items)
        {
            LogHelper.LogFormat("Start populating {0} rows using multi cells and rows method with expression tree", items.Count.ToString());
            var objectArray = ArrayConverter.ConvertTo2DObjectArrayWithExpressionTree<SampleObject>()(items);
            LogHelper.LogFormat("Completed converting the data into string array");

            var dataRange = string.Format("{0}{1}:{2}{3}", ExcelColumnNames.GetColumnName(0), 1, ExcelColumnNames.GetColumnName(objectArray.GetUpperBound(1)), items.Count + 1);
            var r = worksheets[0].get_Range(dataRange);
            r.Value = objectArray;

            var pivotCaches = book.PivotCaches();
            var pivotCache = pivotCaches.Create(XlPivotTableSourceType.xlDatabase, string.Format("RawData!{0}", dataRange));
            //var pivotCache = pivotCaches.Create(XlPivotTableSourceType.xlDatabase, r);

            PivotTables pivotTables = worksheets[1].PivotTables() as PivotTables;
            var pivotTable = pivotTables.Add(pivotCache, worksheets[1].get_Range("A1"), "PivotTable1") as PivotTable;

            var fieldA = pivotTable.PivotFields("A") as PivotField;
            fieldA.Orientation = XlPivotFieldOrientation.xlRowField;

            var fieldB = pivotTable.PivotFields("B") as PivotField;
            fieldB.Orientation = XlPivotFieldOrientation.xlHidden;

            var fieldD = pivotTable.PivotFields("D") as PivotField;
            fieldD.Orientation = XlPivotFieldOrientation.xlPageField;

            var fieldE = pivotTable.PivotFields("E") as PivotField;
            fieldE.Orientation = XlPivotFieldOrientation.xlColumnField;

            var fieldC = pivotTable.PivotFields("C") as PivotField;
            fieldC.Orientation = XlPivotFieldOrientation.xlDataField;
            fieldC.Function = XlConsolidationFunction.xlCount;

            pivotTable.AddDataField(pivotTable.PivotFields("Qty"), "Quantity", XlConsolidationFunction.xlSum);
            pivotTable.AddDataField(pivotTable.PivotFields("Amount"), "SubTotal", XlConsolidationFunction.xlSum);
            pivotTable.AddDataField(pivotTable.PivotFields("Price"), "Price Average", XlConsolidationFunction.xlSum);

            ExcelAPIHelper.DisposeCOMObject(r);
        }

        static void PrintObjectStatically(SampleObject obj)
        {
            Console.WriteLine("PrintObjectStatically");
            Console.WriteLine("A:" + obj.A);
            Console.WriteLine("B:" + obj.B);
            Console.WriteLine("C:" + obj.C);
            Console.WriteLine("D:" + obj.D);
            Console.WriteLine("E:" + obj.E);
        }

        static void PrintObjectWithReflection<T>(T obj)
        {
            Console.WriteLine("PrintObjectWithReflection");
            var t = obj.GetType();

            foreach (var p in t.GetProperties())
            {
                Console.WriteLine(p.Name + ":" + p.GetValue(obj));
            }
        }

        static void PrintObjectWithExpressionTree(SampleObject obj)
        {
            Console.WriteLine("PrintObjectWithExpressionTree");
            var printer = PrintStringValues<SampleObject>();

            printer(obj);
        }

        static Action<T> PrintStringValues<T>()
        {
            var t = typeof(T);

            var statements = new List<Expression>();
            ParameterExpression instanceParameter = Expression.Parameter(t);
            statements.Add(instanceParameter);

            foreach (var p in t.GetProperties())
            {
                var propExpression = Expression.Property(instanceParameter, p);
                MethodCallExpression mCall = Expression.Call(typeof(Console).GetMethod("WriteLine", new Type[] { typeof(string) }), propExpression);
                statements.Add(mCall);
            }

            var body = Expression.Block(statements.ToArray());
            var compiled = Expression.Lambda<Action<T>>(body, instanceParameter).Compile();
            return compiled;
        }

    }
}
