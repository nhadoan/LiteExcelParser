using Excel.Parser;
using System;
using System.Linq;
namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelAgent agent = new ExcelAgent(@"test.xlsx");
            Console.WriteLine("All sheets:");
            var sheets = agent.GetSheets();
            sheets.All(p => {
                Console.WriteLine(p);
                return true;
            });
            Console.WriteLine("-------------");
            foreach (var name in sheets)
            {
                Console.WriteLine("Parsing..." + name);
                agent.Parse(r =>
                {
                    Console.WriteLine("Row " + r.RowIndex);
                    foreach (var c in r.Cells)
                        Console.WriteLine("Column " + c.ColumnIndex + ":" + c.Value);
                }, name);
                Console.WriteLine("-------------");
            }

            Console.WriteLine("Done. Press any key to exit");
            Console.ReadKey();
        }
    }
}
