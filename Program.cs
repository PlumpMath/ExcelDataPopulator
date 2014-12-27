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
    class SampleObject
    {
        public string A { get; set; }
        public string B { get; set; }
        public string C { get; set; }
        public string D { get; set; }
        public string E { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {
            var tempFilePath = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(Path.GetRandomFileName(), ".xlsx")); ;

            var items = Enumerable.Range(1, 10000).Select(i => new SampleObject { A = "A" + i, B = "B" + i, C = "C" + i, D = "D" + i, E = "E" + i }).ToList();
            //Enumerable.Range(0, 127).Aggregate((x, y) => { Debug.WriteLine(y + " = " + (char)y); return 0; });

            //Populate(tempFilePath, WaysToAssignValues);
            Populate(tempFilePath, ws => PopulateItemsWithSingleCellMethod(ws, items));

            //Process.Start(tempFilePath);
            //Console.Read();

            // clean up
            //File.Delete(filePath);
        }

        // example of ways to set cell value 
        static void WaysToAssignValues(Worksheet ws)
        {
            // DO NOT do this as it will create memory leak (excel stays as background process, visible in task manager)
            // best pratice is to hold the object into a variable then call dispose (release com object)
            //ws.get_Range("C5").Value = "ABC";

            // single cell
            var r1 = ws.get_Range("A1");
            r1.Value = "A1";
            DisposeCOMObject(r1);

            var r2 = ws.get_Range("A2:C2"); // multiple cell in 1 row
            r2.Value = new[] { "A2", "B2", "C2" };
            DisposeCOMObject(r2);

            var r3 = ws.get_Range("A3:C4"); // multiple cells/rows
            r3.Value = new string[,] { { "A3", "B3", "C3" }, { "A4", "B4", "C4" } }; ;
            DisposeCOMObject(r3);
        }

        static void PopulateItemsWithSingleCellMethod(Worksheet ws, IList<SampleObject> items)
        {
            LogFormat("Start populating {0} rows using single cell method", items.Count.ToString());
            for (int i = 0; i < items.Count; i++)
            {
                var item = items[i];
                var rowIndex = i + 1;
                var r1 = ws.get_Range("A" + rowIndex); r1.Value = item.A; DisposeCOMObject(r1);
                var r2 = ws.get_Range("B" + rowIndex); r2.Value = item.B; DisposeCOMObject(r2);
                var r3 = ws.get_Range("C" + rowIndex); r3.Value = item.C; DisposeCOMObject(r3);
                var r4 = ws.get_Range("D" + rowIndex); r4.Value = item.D; DisposeCOMObject(r4);
                var r5 = ws.get_Range("E" + rowIndex); r5.Value = item.E; DisposeCOMObject(r5);
            }
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
