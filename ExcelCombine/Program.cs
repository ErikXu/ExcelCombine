using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelCombine
{
    class Program
    {
        static void Main()
        {
            var folder = ConfigurationManager.AppSettings["Folder"];
            var firstRowIndex = int.Parse(ConfigurationManager.AppSettings["FirstRowIndex"]);

            var files = Directory.GetFiles(folder)
                .Where(n => Path.GetExtension(n).Equals(".xlsx", StringComparison.OrdinalIgnoreCase)).OrderBy(n => n)
                .ToList();

            if (files.Count < 2)
            {
                Console.WriteLine("The folder should contain more than one excel files with extension .xlsx.");
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
                return;
            }

            var resultFolder = Path.Combine(folder, "Result");
            if (!Directory.Exists(resultFolder))
            {
                Directory.CreateDirectory(resultFolder);
            }

            var filePath = Path.Combine(resultFolder, "Combine.xlsx");

            using (var firstStream = new FileStream(files[0], FileMode.Open, FileAccess.Read))
            {
                using (var stream = new FileStream(filePath, FileMode.Create))
                {
                    var combinedWorkbook = new XSSFWorkbook(firstStream);

                    var sheetCount = combinedWorkbook.NumberOfSheets;
                    if (sheetCount > 1)
                    {
                        for (int i = 1; i < sheetCount; i++)
                        {
                            combinedWorkbook.RemoveSheetAt(i);
                        }
                    }

                    var combinedSheet = combinedWorkbook.GetSheetAt(0);

                    var dataFormat = combinedWorkbook.CreateDataFormat();
                    var dateStyle = combinedWorkbook.CreateCellStyle();
                    dateStyle.DataFormat = dataFormat.GetFormat("yyyy/MM/dd HH:mm:ss");

                    CombineFiles(files, firstRowIndex, combinedSheet, combinedSheet.LastRowNum, dateStyle);
                    combinedWorkbook.Write(stream);
                }
            }

            Console.WriteLine("Combine success.");
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();
        }

        private static void CombineFiles(List<string> files, int firstRowIndex, ISheet combinedSheet, int lastRowNum, ICellStyle dateStyle)
        {
            for (var i = 1; i < files.Count; i++)
            {
                using (var readStream = new FileStream(files[i], FileMode.Open, FileAccess.Read))
                {
                    var workbook = new XSSFWorkbook(readStream);
                    var sheet = workbook.GetSheetAt(0);

                    var evaluator = new XSSFFormulaEvaluator(workbook);

                    for (var j = firstRowIndex - 1; j <= sheet.LastRowNum; j++)
                    {
                        var combinedRow = combinedSheet.CreateRow(++lastRowNum);
                        var row = sheet.GetRow(j);
                        if (row == null)
                        {
                            continue;
                        }

                        var cellCount = row.Cells.Count;
                        if (cellCount <= 0)
                        {
                            continue;
                        }

                        foreach (var cell in row.Cells)
                        {
                            var combinedCell = combinedRow.CreateCell(cell.ColumnIndex);
                            CopyCell(cell, combinedCell, evaluator, dateStyle);
                        }
                    }
                }

                Console.WriteLine("[{0}/{1}]{2} is combined.", i + 1, files.Count, files[i]);
            }
        }

        private static void CopyCell(ICell srcCell, ICell tarCell, IFormulaEvaluator evaluator, ICellStyle dateStyle)
        {
            tarCell.SetCellType(srcCell.CellType);

            switch (srcCell.CellType)
            {
                case CellType.Blank:
                    tarCell.SetCellValue(srcCell.StringCellValue);
                    break;
                case CellType.Boolean:
                    tarCell.SetCellValue(srcCell.BooleanCellValue);
                    break;
                case CellType.Error:
                    tarCell.SetCellValue(srcCell.ErrorCellValue);
                    break;
                case CellType.Formula:
                    var cell = evaluator.EvaluateInCell(srcCell);
                    tarCell.SetCellType(CellType.String);
                    tarCell.SetCellValue(cell.ToString());
                    break;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(srcCell))
                    {
                        tarCell.CellStyle = dateStyle;
                        tarCell.SetCellValue(srcCell.DateCellValue);
                    }
                    else
                    {
                        tarCell.SetCellValue(srcCell.NumericCellValue);
                    }
                    break;
                case CellType.String:
                    tarCell.SetCellValue(srcCell.StringCellValue);
                    break;
                case CellType.Unknown:
                    tarCell.SetCellValue(srcCell.StringCellValue);
                    break;
            }
        }
    }
}
