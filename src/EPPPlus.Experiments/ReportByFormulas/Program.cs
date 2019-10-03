using OfficeOpenXml;
using System;
using System.IO;

namespace ReportByFormulas
{
    class Program
    {
        static void Main(string[] args)
        {

            //Loads or open an existing workbook through Open method of IWorkbooks
            var inputFileName = @"Template.xlsx";
            var resultFileName = @"Report.xlsx";

            var fileTemplate = new FileInfo(inputFileName);
            var fileResult = new FileInfo(resultFileName);
            using (var excelPackage = new ExcelPackage(fileTemplate))
            {
                var ws = excelPackage.Workbook.Worksheets["Data"];

                // Read number value
                var numberCell = ws.Cells["D2"];
                PrintCellInfo(numberCell, "Number value");

                // Read date value
                var dateCell = ws.Cells[2, 1];
                PrintCellInfo(dateCell, "Date value");

                // Read string value
                var strCell = ws.Cells[2, 3];
                PrintCellInfo(strCell, "String value");

                // Change cell value
                var cellToChange = ws.Cells["D3"];
                cellToChange.Value = 500.23;

                ws.Calculate();

                // Read formula and value
                Console.WriteLine("Formula value");
                var formulaCell = ws.Cells[1, 7];
                Console.WriteLine($"value: {formulaCell.Value}, text: {formulaCell.Text}, formula: {formulaCell.Formula}");

                var expectedValue = 10700.46;
                Console.WriteLine($"Formula value: {formulaCell.Value}, Expected value: {expectedValue}");

                // Save workbook
                excelPackage.SaveAs(fileResult);

            }

            Console.WriteLine("Press any key to finish...");
            Console.ReadKey();
        }

        /// <summary>
        /// Prints info about cell to console.
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="title"></param>
        private static void PrintCellInfo(ExcelRange cell, string title)
        {
            Console.WriteLine(title);
            Console.WriteLine($"value: {cell.Value}, value type: {cell.Value.GetType()}, text: {cell.Text}, format: {cell.Style.Numberformat.NumFmtID}");

        }
    }
}
