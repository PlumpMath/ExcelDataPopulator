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
    class ExcelAPIHelper
    {
        public static string CreateNew(int numberOfWorksheets, Action<Workbook, IList<Worksheet>> populateData)
        {
            var filePath = Path.Combine(Path.GetTempPath(), Path.ChangeExtension(Path.GetRandomFileName(), ".xlsx")); ;

            Application xlApp = null;
            Workbooks newBooks = null;
            Workbook newBook = null;
            Sheets newBookWorksheets = null;
            //Worksheet defaultWorksheet = null;

            IList<Worksheet> worksheetsToReturn = new List<Worksheet>();

            try
            {
                LogHelper.LogFormat("Starting excel");
                xlApp = new Application { Visible = false, DisplayAlerts = false };

                if (xlApp == null)
                {
                    LogHelper.LogFormat("EXCEL could not be started. Check that your office installation and project references are correct.");
                    return null;
                }

                //LogHelper.LogFormat("Create a new workbook, comes with an empty default worksheet");
                newBooks = xlApp.Workbooks;
                newBook = newBooks.Add(XlWBATemplate.xlWBATWorksheet);
                newBookWorksheets = newBook.Worksheets;

                // GET THE REFERENCE FOR THE EMPTY DEFAULT WORKSHEET
                if (newBookWorksheets.Count > 0)
                {
                    var defaultWs = newBookWorksheets[1] as Worksheet;
                    defaultWs.Name = "RawData";
                    worksheetsToReturn.Add(defaultWs);
                }

                if (numberOfWorksheets > 1)
                {
                    for (int i = 0; i < numberOfWorksheets - 1; i++)
                    {
                        var newWS = newBookWorksheets.Add() as Worksheet;
                        worksheetsToReturn.Add(newWS);
                    }
                }

                var sw = Stopwatch.StartNew();
                populateData(newBook, worksheetsToReturn);
                LogHelper.LogFormat("Populate data completed in {0} ms", sw.ElapsedMilliseconds.ToString());

                LogHelper.LogFormat("Saving the new book into the export file path: {0}", filePath);
                newBook.SaveAs(filePath);

                newBooks.Close();

                return filePath;
            }
            catch (Exception ex)
            {
                LogHelper.LogFormat("Method: Populate - Exception: {0}", ex.ToString());
                return null;
            }
            finally
            {
                foreach (var ws in worksheetsToReturn)
                {
                    DisposeCOMObject(ws);    
                }
                
                DisposeCOMObject(newBookWorksheets);
                DisposeCOMObject(newBook);
                DisposeCOMObject(newBooks);


                LogHelper.LogFormat("Closing the excel app");
                if (xlApp != null)
                {
                    xlApp.Quit();
                    DisposeCOMObject(xlApp);
                }
            }
        }

        public static void DisposeCOMObject(object o)
        {
            //LogHelper.LogFormat("Method: DisposeCOMObject - Disposing");
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
                LogHelper.LogFormat("Method: DisposeCOMObject - Exception: {0}", ex.ToString());
            }
        }
    }
}
