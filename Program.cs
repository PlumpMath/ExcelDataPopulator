using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelDataPopulator
{
    class Program
    {
        static void Main(string[] args)
        {
            var sw = Stopwatch.StartNew();
            //new NumberGenerator().Generate();

            RunPopulate();

            Console.WriteLine("Finished in {0} ms", sw.ElapsedMilliseconds);
            sw.Stop();
            Console.Read();
        }

        static void RunPopulate()
        {
            var tempFilePath = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(Path.GetRandomFileName(), ".xlsx")); ;

            // prepare data source
            var items = Enumerable.Range(1, 100000).Select(i => new SampleObject { A = "A" + i, B = "B" + i, C = "C" + i, D = "D" + i, E = "E" + i }).ToList();

            // uncomment 1 option below to verify the performance and results
            //Populate(tempFilePath, WaysToAssignValues);
            //Populate(tempFilePath, ws => PopulateItemsWithSingleCellMethod(ws, items));
            //Populate(tempFilePath, ws => PopulateItemsWithMultiCellsOneRowMethod(ws, items));
            Populate(tempFilePath, ws => PopulateItemsWithMultiCellsAndRowsMethod(ws, items));

            Process.Start(tempFilePath);
            Console.Read();

            // clean up
            //File.Delete(filePath);
        }

        static void WaysToAssignValues(Worksheet worksheet)
        {
            LogFormat("Ways to populate values into excel");
            // DO NOT do this as it will create memory leak (excel stays as background process, visible in task manager)
            // best pratice is to hold the object into a variable then call dispose (release com object)
            //ws.get_Range("C5").Value = "ABC";

            var r1 = worksheet.get_Range("A1"); // single cell
            r1.Value = "A1";
            DisposeCOMObject(r1);

            var r2 = worksheet.get_Range("A2:C2"); // multi cells in 1 row
            r2.Value = new[] { "A2", "B2", "C2" };
            DisposeCOMObject(r2);

            var r3 = worksheet.get_Range("A3:C4"); // multi cells/rows
            r3.Value = new string[,] { { "A3", "B3", "C3" }, { "A4", "B4", "C4" } }; ;
            DisposeCOMObject(r3);
        }

        static void PopulateItemsWithSingleCellMethod(Worksheet worksheet, IList<SampleObject> items)
        {
            LogFormat("Start populating {0} rows using single cell method", items.Count.ToString());
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var rowIndex = i + 1;
                var r1 = worksheet.get_Range("A" + rowIndex); r1.Value = item.A; DisposeCOMObject(r1);
                var r2 = worksheet.get_Range("B" + rowIndex); r2.Value = item.B; DisposeCOMObject(r2);
                var r3 = worksheet.get_Range("C" + rowIndex); r3.Value = item.C; DisposeCOMObject(r3);
                var r4 = worksheet.get_Range("D" + rowIndex); r4.Value = item.D; DisposeCOMObject(r4);
                var r5 = worksheet.get_Range("E" + rowIndex); r5.Value = item.E; DisposeCOMObject(r5);
            }
        }

        static void PopulateItemsWithMultiCellsOneRowMethod(Worksheet worksheet, IList<SampleObject> items)
        {
            LogFormat("Start populating {0} rows using multi cells (per 1 row) method", items.Count.ToString());
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var rowIndex = i + 1;
                var r = worksheet.get_Range(string.Format("A{0}:E{0}", rowIndex));
                r.Value = new[] { item.A, item.B, item.C, item.D, item.E };
                DisposeCOMObject(r);
            }
        }

        static void PopulateItemsWithMultiCellsAndRowsMethod(Worksheet worksheet, IList<SampleObject> items)
        {
            LogFormat("Start populating {0} rows using multi cells and rows method", items.Count.ToString());
            var stringArray = new string[items.Count, 5];
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var rowIndex = i + 1;
                stringArray[i, 0] = item.A;
                stringArray[i, 1] = item.B;
                stringArray[i, 2] = item.C;
                stringArray[i, 3] = item.D;
                stringArray[i, 4] = item.E;
            }

            var r = worksheet.get_Range(string.Format("A{0}:E{1}", 1, items.Count));
            r.Value = stringArray;
            DisposeCOMObject(r);
        }

        static void Populate(string filePath, Action<Worksheet> populateData)
        {
            Application xlApp = null;
            Workbooks newBooks = null;
            Workbook newBook = null;
            Sheets newBookWorksheets = null;
            Worksheet defaultWorksheet = null;

            try
            {
                LogFormat("Starting excel");
                xlApp = new Application { Visible = false, DisplayAlerts = false };

                if (xlApp == null)
                {
                    LogFormat("EXCEL could not be started. Check that your office installation and project references are correct.");
                    return;
                }

                //LogFormat("Create a new workbook, comes with an empty default worksheet");
                newBooks = xlApp.Workbooks;
                newBook = newBooks.Add(XlWBATemplate.xlWBATWorksheet);
                newBookWorksheets = newBook.Worksheets;

                // GET THE REFERENCE FOR THE EMPTY DEFAULT WORKSHEET
                if (newBookWorksheets.Count > 0)
                {
                    defaultWorksheet = newBookWorksheets[1] as Worksheet;
                }

                var sw = Stopwatch.StartNew();
                populateData(defaultWorksheet);
                LogFormat("Populate data completed in {0} ms", sw.ElapsedMilliseconds.ToString());

                LogFormat("Saving the new book into the export file path: {0}", filePath);
                newBook.SaveAs(filePath);

                newBooks.Close();
            }
            catch (Exception ex)
            {
                LogFormat("Method: Populate - Exception: {0}", ex.ToString());
            }
            finally
            {
                DisposeCOMObject(defaultWorksheet);
                DisposeCOMObject(newBookWorksheets);
                DisposeCOMObject(newBook);
                DisposeCOMObject(newBooks);


                LogFormat("Closing the excel app");
                if (xlApp != null)
                {
                    xlApp.Quit();
                    DisposeCOMObject(xlApp);
                }
            }
        }

        static void DisposeCOMObject(object o)
        {
            //LogFormat("Method: DisposeCOMObject - Disposing");
            if (o == null)
            {
                return;
            }
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            catch (Exception ex)
            {
                LogFormat("Method: DisposeCOMObject - Exception: {0}", ex.ToString());
            }
        }

        static void LogFormat(string format, params string[] message)
        {
            Console.WriteLine(string.Format("{0} - {1}", DateTime.Now.ToString("u"), string.Format(format, message)));
        }

    }


}
