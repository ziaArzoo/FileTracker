using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Threading.Tasks;


class ExcelUpdater
{
    static string PickFile(string title, int num)
    {
        using (var dialog = new OpenFileDialog())
        {
            dialog.Title = title;
            dialog.Filter = "Excel Files|*.xlsx;*.xls";
            dialog.Multiselect = false;
            DialogResult result = dialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                Console.Write($"reading files{num}");
                for (int k = 0; k < 5; k++)
                {
                    Console.Write(".");
                    Thread.Sleep(500);  // it's fine here
                }
                Console.Write($"\nfiles{num} Read\n");
                return dialog.FileName;
            }
            else
            {
                throw new Exception("File selection cancelled.");
            }
        }
    }
    [STAThread]
    static void Main(string[] args)
    {
        //making the UI
        string inputFile1 = null;
        string inputFile2 = null;
        Console.ForegroundColor = ConsoleColor.Cyan;
        Console.WriteLine("\n=========================================");
        Console.WriteLine("             Excel Merge Tool v1.0         ");
        Console.WriteLine("=========================================\n");
        Console.ResetColor();

        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.Write("Hi");
        int j = 0;
        while (j < 3)
        {
           Console.Write("!");
            Thread.Sleep(300);  // waits 300ms second before printing next dot
            j++;
        }
        Console.WriteLine("\nDo you want Me to import files(Yes/No)");
        string input = Console.ReadLine();
        //Prompting User for a value
        if (input.ToLower() == "yes")
        {
            inputFile1 = PickFile("Select the first Excel file where", 1);
            inputFile2 = PickFile("Select the first Excel file where", 2);
            try
            {
                // Validate both inputs
                if (!File.Exists(inputFile1) || !File.Exists(inputFile2))
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("One or both input files are invalid. Please check the paths.");
                    Console.ResetColor();
                }

                // Output folder next to the first input
                string baseDir = Path.GetDirectoryName(inputFile1);
                string outputDir = Path.Combine(baseDir, "Merge_updated_file");
                Directory.CreateDirectory(outputDir);

                // Output file name
                string outputFileName = "Combined" + ".xlsx";
                string outputPath = Path.Combine(outputDir, outputFileName);
                string keyColumn = "Id";


                var file1Data = ReadExcelToDict(inputFile1, keyColumn, out var headers);
                var file2Data = ReadExcelToDict(inputFile2, keyColumn, out _);

                Console.ForegroundColor = ConsoleColor.Blue;
                Console.Write("\n Merging data");
                for (int i = 0; i < 5; i++) { Thread.Sleep(200); Console.Write("."); }
                Console.ResetColor();

                foreach (var kvp in file2Data)
                {
                    string id = kvp.Key;
                    var newRow = kvp.Value;

                    if (file1Data.ContainsKey(id))
                    {
                        var oldRow = file1Data[id];
                        foreach (var col in newRow.Keys)
                        {
                            if (oldRow.ContainsKey(col) && oldRow[col] != newRow[col])
                                oldRow[col] = newRow[col];
                        }
                    }
                    else
                    {
                        file1Data[id] = newRow;
                    }
                }

                WriteExcelFromDict(outputPath, headers, file1Data);

                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"\n Success! Merged file saved to:\n{outputPath}");
                Console.ResetColor();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"\n Error: {ex.Message}");
                Console.ResetColor();
            }
            finally
            {
                Console.ForegroundColor = ConsoleColor.DarkGray;
                Console.WriteLine("\nPress any key to exit...");
                Console.ResetColor();
                Console.ReadKey();
            }

        }
        else
        {
            Console.WriteLine("Thank see you later!!");
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine("\nPress any key to exit...");
            Console.ResetColor();
            Console.ReadKey();
        }
    }
    // ↓↓↓ Other methods stay unchanged ↓↓↓

    static Dictionary<string, Dictionary<string, string>> ReadExcelToDict(string path, string keyColumn, out List<string> headers)
    {
        var result = new Dictionary<string, Dictionary<string, string>>();
        headers = new List<string>();

        using (var doc = SpreadsheetDocument.Open(path, false))
        {
            var sheet = doc.WorkbookPart.Workbook.Sheets.Elements<Sheet>().First();
            var wsPart = (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
            var rows = wsPart.Worksheet.Descendants<Row>().ToList();
            var sharedStrings = doc.WorkbookPart.SharedStringTablePart?.SharedStringTable;

            if (rows.Count == 0)
                throw new Exception("No rows found in Excel file.");

            var headerCells = rows[0].Elements<Cell>().ToList();
            var headerDict = new Dictionary<string, string>();

            foreach (var cell in headerCells)
            {
                string col = GetColumnName(cell.CellReference);
                string val = GetCellValue(cell, sharedStrings);
                headerDict[col] = val;
            }

            headers = headerDict.OrderBy(k => ExcelColumnToIndex(k.Key)).Select(kv => kv.Value).ToList();

            int keyIndex = headers.IndexOf(keyColumn);
            if (keyIndex == -1)
                throw new Exception("Key column not found in headers.");

            foreach (var row in rows.Skip(1))
            {
                var rowDict = new Dictionary<string, string>();
                var cellDict = row.Elements<Cell>().ToDictionary(
                    c => GetColumnName(c.CellReference),
                    c => c);

                for (int i = 0; i < headers.Count; i++)
                {
                    string col = GetExcelColumnName(i + 1);
                    cellDict.TryGetValue(col, out var cell);
                    rowDict[headers[i]] = GetCellValue(cell, sharedStrings);
                }

                string key = rowDict[keyColumn];
                if (!string.IsNullOrWhiteSpace(key))
                    result[key] = rowDict;
            }
        }

        return result;
    }

    static void WriteExcelFromDict(string path, List<string> headers, Dictionary<string, Dictionary<string, string>> data)
    {
        using (SpreadsheetDocument doc = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
        {
            var workbookPart = doc.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var wsPart = workbookPart.AddNewPart<WorksheetPart>();
            var sheetData = new SheetData();
            wsPart.Worksheet = new Worksheet(sheetData);

            var headerRow = new Row();
            foreach (var header in headers)
                headerRow.Append(CreateCell(header));
            sheetData.Append(headerRow);

            foreach (var rowDict in data.Values)
            {
                var row = new Row();
                foreach (var header in headers)
                {
                    rowDict.TryGetValue(header, out string val);
                    row.Append(CreateCell(val ?? ""));
                }
                sheetData.Append(row);
            }

            var sheets = doc.WorkbookPart.Workbook.AppendChild(new Sheets());
            var sheet = new Sheet()
            {
                Id = doc.WorkbookPart.GetIdOfPart(wsPart),
                SheetId = 1,
                Name = "Merged"
            };
            sheets.Append(sheet);
        }
    }

    static Cell CreateCell(string text)
    {
        if (double.TryParse(text, out _))
            return new Cell { DataType = CellValues.Number, CellValue = new CellValue(text) };

        if (text == "TRUE" || text == "FALSE")
            return new Cell { DataType = CellValues.Boolean, CellValue = new CellValue(text == "TRUE" ? "1" : "0") };

        return new Cell { DataType = CellValues.String, CellValue = new CellValue(text) };
    }

    static string GetCellValue(Cell cell, SharedStringTable sharedStrings)
    {
        if (cell == null || cell.CellValue == null)
            return "";

        string val = cell.CellValue.InnerText;

        if (cell.DataType != null && cell.DataType == CellValues.SharedString)
        {
            if (int.TryParse(val, out int idx) && sharedStrings != null)
                return sharedStrings.ElementAt(idx).InnerText;
        }

        return val;
    }

    static string GetColumnName(string cellRef)
    {
        if (string.IsNullOrWhiteSpace(cellRef)) return "";
        return Regex.Match(cellRef, @"[A-Za-z]+").Value;
    }

    static int ExcelColumnToIndex(string column)
    {
        int index = 0;
        foreach (char c in column)
        {
            index *= 26;
            index += (c - 'A' + 1);
        }
        return index;
    }

    static string GetExcelColumnName(int index)
    {
        string column = "";
        while (index > 0)
        {
            index--;
            column = (char)('A' + (index % 26)) + column;
            index /= 26;
        }
        return column;
    }
}
