using Excel.Parser;
using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace ConsoleApp
{
    class Program
    {

        static void Main(string[] args)
        {
            ExcelAgent agent = new ExcelAgent(@"SampleData.xlsx");
            Console.WriteLine("All sheets:");
            var sheets = agent.GetSheets();
            sheets.All(p => {
                Console.WriteLine(p);
                return true;
            });
            Console.WriteLine("-------------");

            Regex sheetExp = new Regex(@"^([a-z]{3} \d{4})$", RegexOptions.Compiled | RegexOptions.IgnoreCase);
            string[] validRows = new string[] { "App", "Ticket machine" };
            foreach (var name in sheets)
            {
                if (sheetExp.IsMatch(name))
                {
                    Console.WriteLine("Parsing..." + name);
                    ExcelAgent.ParsedRow firstRow = null;
                    agent.Parse(r =>
                    {

                        if (r.RowIndex == 1)
                        {
                            firstRow = r;
                        }
                        else
                        {
                            
                            var firstCol = r.Cells[0];
                            if (firstCol.ColumnIndex == "A" && validRows.Contains(firstCol.Value))
                            {
                                Console.WriteLine("Row " + r.RowIndex + ": " + firstCol.Value);
                                foreach (var c in firstRow.Cells) if ("BCDEFGHI".IndexOf(c.ColumnIndex) >= 0) Console.WriteLine("\t" + firstRow[c.ColumnIndex].Value + ":" + r[c.ColumnIndex]?.Value + "; ");
                                Console.WriteLine();
                            }
                        }
                    }, name);
                    Console.WriteLine("-------------");
                }
                
            }

            Console.WriteLine("Done. Press any key to exit");
            Console.ReadKey();
        }

        static void MainBasic(string[] args)
        {
            ExcelAgent agent = new ExcelAgent(@"SampleData.xlsx");
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
