using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Excel.Parser
{
    public class ExcelAgent
    {
        string _file = null;
        public ExcelAgent(string file)
        {
            _file = file;
        }
        static string GetSharedStringCellValue(WorkbookPart workbookPart, Cell cell)
        {
            SharedStringTablePart sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            SharedStringTable ssTable = sstPart.SharedStringTable;
            return ssTable.ChildElements[Convert.ToInt32(cell.CellValue.Text)].InnerText;

        }

        public string[] GetSheets()
        {
            List<string> names = new List<string>();
            using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(_file, false))
            {
                foreach(Sheet sh in myDoc.WorkbookPart.Workbook.Sheets) names.Add(sh.Name);                
                myDoc.Close();
            }
            return names.ToArray();
        }
        public async Task ParseAsync(Action<ParsedRow> action, string sheetName = null)
        {
            await Task.Run(() =>
            {
                this.Parse(action, sheetName);
            });
        }
        public void Parse(Action<ParsedRow> action, string sheetName = null)
        {
            
            using (SpreadsheetDocument myDoc = SpreadsheetDocument.Open(_file, false))
            {

                WorkbookPart workbookPart = myDoc.WorkbookPart;

                
                
                WorksheetPart worksheet = null;
                if (sheetName == null)
                {
                    Sheet sheet = (Sheet)myDoc.WorkbookPart.Workbook.Sheets.First();
                    sheetName = sheet.Name;
                    worksheet = workbookPart.WorksheetParts.First();
                }
                else
                {
                    Sheet sheet = (Sheet)myDoc.WorkbookPart.Workbook.Sheets.First(sh => ((Sheet)sh).Name.Value.ToLower() == sheetName.ToLower());
                    if(sheet!=null) worksheet = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                }

                if (worksheet == null) throw new Exception("Can not find worksheet named: " + sheetName);

                OpenXmlReader reader = OpenXmlReader.Create(worksheet);
                while (reader.Read())
                {
                    if (reader.IsStartElement)
                    {
                        if (reader.ElementType == typeof(Row))
                        {
                            Row row = (Row)reader.LoadCurrentElement();
                            IEnumerable<Cell> cells = row.Elements<Cell>();
                            List<Cell> orgCells = new List<Cell>();
                            foreach (var cell in cells)
                            {
                                orgCells.Add(cell);                                
                            }
                            action(new ParsedRow((int)row.RowIndex.Value, orgCells.ToArray(), sheetName, workbookPart));
                        }
                       
                    }
                }
                reader.Close();
                myDoc.Close();
            }
        }
        public class ParsedRow
        {
            public string SheetName { get; private set; }
            public int RowIndex { get; private set; }
            public ParsedCell[] Cells { get; private set; }
            public Cell[] OrgCells { get; private set; }
            internal ParsedRow(int index, Cell[] cells, string sheetName, WorkbookPart wp)
            {
                RowIndex = index;
                OrgCells = cells;
                SheetName = sheetName;
                //parser cell;
                ParseCells(wp);
            }
            protected void ParseCells(WorkbookPart wp)
            {
                Cells = new ParsedCell[OrgCells.Length];
                for (int i=0;i<Cells.Length;i++)
                {
                    var cell = OrgCells[i];
                    string cellValue = "";
                    if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                    {
                        cellValue = GetSharedStringCellValue(wp, cell);
                    }
                    else
                    {
                        cellValue = cell.InnerText;
                    }

                    Cells[i] = new ParsedCell(cell.CellReference.ToString().Replace(RowIndex.ToString(), ""), cellValue);
                }
            }
        }
        public class ParsedCell
        {
            public string ColumnIndex { get; private set; }
            public string Value { get; private set; }
            internal ParsedCell(string index, string value)
            {
                ColumnIndex = index;
                Value = value;
            }
        }
        
    }
}
