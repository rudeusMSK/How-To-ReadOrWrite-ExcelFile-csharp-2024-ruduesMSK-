using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;

namespace ReadExcelFile_v._2
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // in tiếng việt ra Console c#
            Console.OutputEncoding = Encoding.UTF8;

            // Đường dẫn tới file Excel
            string fileName = @"..\..\File\excelFile.xlsx";

            // Đọc bình thường
            Console.WriteLine("Đọc bình thường");
            ReadExcel(fileName);

            Console.WriteLine("\n");

            // Đọc theo từng ô
            Console.WriteLine("đọc theo từng ô");
            ReadExcelv2(fileName);

            Console.WriteLine();

            // Đọc theo ô cố định
            Console.WriteLine("đọc theo ô cố định");
            string C16 = ReadCell(fileName, "C16");
            string D16 = ReadCell(fileName, "D16");
            string E16 = ReadCell(fileName, "E16");
            Console.WriteLine($"C16: {C16}");
            Console.WriteLine($"D16: {D16}");
            Console.WriteLine($"E16: {E16}");

            Console.WriteLine();

            // Đọc dòng cuối, nhưng ko biết 9 xác ở dòng nào !
            Console.WriteLine("Đọc dòng cuối");
            ReadLastRow(fileName);

            Console.ReadKey();
        }
        static void ReadLastRow(string fileName)
        {
            // Mở file Excel
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                // Lưu trữ giá trị của dòng cuối cùng
                string[] lastRowValues = null;

                // Duyệt qua tất cả các dòng trong sheet
                foreach (Row row in sheetData.Elements<Row>())
                {
                    // Lưu trữ giá trị của dòng hiện tại
                    lastRowValues = row.Elements<Cell>().Select(cell => GetValueOfCell(workbookPart, cell)).ToArray();
                }

                // Kiểm tra xem có dòng nào trong sheet không
                if (lastRowValues != null)
                {
                    Console.WriteLine("Dòng cuối nè:");
                    foreach (var value in lastRowValues)
                    {
                        Console.Write(value + " ");
                    }
                }
                else Console.WriteLine("Excel chưa có zều hết ahihi");
            }
        }

        static void ReadExcel(string fileName)
        {
            // Mở file Excel
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                // Duyệt qua các ô trong sheet và hiển thị giá trị
                foreach (Row row in sheetData.Elements<Row>())
                {
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        Console.Write(GetValueOfCell(workbookPart, cell) + " ");
                    }
                }
            }
        }

        static void ReadExcelv2(string fileName)
        {
            // Mở file Excel
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                // Duyệt qua các ô trong sheet và hiển thị giá trị
                foreach (Row row in sheetData.Elements<Row>())
                {
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        // Lấy tọa độ của ô
                        string cellReference = cell.CellReference;

                        // Lấy giá trị của ô
                        string cellValue = GetValueOfCell(workbookPart, cell);

                        // Hiển thị thông tin của ô
                        Console.WriteLine($"vị trí: {cellReference}: {cellValue}");
                    }
                }
            }
        }

        static string ReadCell(string fileName, string cellReference)
        {
            // Mở file Excel
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                // Tách tọa độ cột và hàng từ tên ô
                string columnName = cellReference.Substring(0, cellReference.Length - 1);
                string rowNumber = cellReference.Substring(cellReference.Length - 1);

                // Tìm ô trong sheet
                Cell cell = sheetData.Descendants<Cell>()
                                      .Where(c => string.Compare(c.CellReference.Value, cellReference, true) == 0)
                                      .FirstOrDefault();

                // Nếu không tìm thấy ô, trả về chuỗi rỗng
                if (cell == null)
                    return "";

                // Lấy giá trị của ô
                string value = GetValueOfCell(workbookPart, cell);

                return value;
            }
        }

        static string GetValueOfCell(WorkbookPart workbookPart, Cell cell)
        {
            // Lấy giá trị của ô từ phần CellValue hoặc SharedStringTable
            string value = cell.CellValue?.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (stringTable != null)
                {
                    value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }

            return value ?? string.Empty;
        }
    }
}
