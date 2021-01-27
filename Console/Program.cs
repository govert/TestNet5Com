using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace TestNet5Com
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            DoMyExcelStuffAndCleanup();
            Console.WriteLine("All Done");
        }
        static void DoMyExcelStuffAndCleanup()
        {
            DoMyExcelStuff();

            // Now let the GC clean up (repeat, until no more)
            do
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            while (Marshal.AreComObjectsAvailableForCleanup());
        }

        static void DoMyExcelStuff()
        {
            Application app = new Application();
            app.Visible = true;
            var wb = app.Workbooks.Add();
            var ws = wb.Worksheets[1] as Worksheet;
            ws.Range["A1"].Value = "Hello";
            ws.Range["A2"].Formula = "=2+2";
            ws.Calculate();
            var result = ws.Range["A2"].Formula;
            wb.Close(false);
            Console.WriteLine($"Result: {result}");
            app.Quit();
        }
    }


}
