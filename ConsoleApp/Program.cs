using Excel.Parser;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ConsoleApp
{
    class Program
    {
        class HLTSalesHistory
        {
            public string YearMonth { get; set; }
            public string TicketType { get; set; }
            public string SaleChannel { get; set; }
            public string TicketCount { get; set; }
            public string SaleAmount { get; set; }
        }
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
            List<HLTSalesHistory> finalList = new List<HLTSalesHistory>();
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
                                foreach (var c in firstRow.Cells) 
                                    if ("BCDEFGHI".IndexOf(c.ColumnIndex) >= 0)
                                    {
                                        var pp = finalList.FirstOrDefault(p => p.YearMonth == name && p.TicketType == firstCol.Value && p.SaleChannel == firstRow[c.ColumnIndex].Value);
                                        if ( pp!= null)
                                        {
                                            pp.TicketCount = r[c.ColumnIndex]?.Value;
                                        }
                                        else finalList.Add(new HLTSalesHistory()
                                        {
                                            YearMonth = name,
                                            TicketType = firstCol.Value,
                                            SaleChannel = firstRow[c.ColumnIndex].Value,
                                            SaleAmount = r[c.ColumnIndex]?.Value
                                        });
                                        //Console.WriteLine("\t" + name + "-" + firstCol.Value + firstRow[c.ColumnIndex].Value + r[c.ColumnIndex]?.Value);
                                    }
                            }
                        }
                    }, name);
                    Console.WriteLine("-------------");
                }
                
            }
            for(int i = 0; i < finalList.Count; i++)
            {
                Console.WriteLine(string.Format("{0}\t{1}\t{2}\t{3}\t{4}", finalList[i].YearMonth, finalList[i].TicketType, finalList[i].SaleChannel, finalList[i].SaleAmount, finalList[i].TicketCount));
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
