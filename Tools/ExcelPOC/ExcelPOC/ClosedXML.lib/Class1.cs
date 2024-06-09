using ClosedXML.Excel;
using System;

namespace ClosedXML.Lib
{
    public static class Test
    {
        public static void Main()
        {
            string prefixPath = @"D:\ARC\Code\Sitecore\GIT\sheet-editor\Tools\ExcelPOC\ExcelPOC\ClosedXML.lib\";

            string filePath = prefixPath+"ClosedXmlExample.xlsx";

            // Create and style the Excel file
            CreateAndStyleExcelFile(filePath);

            // Read and display the data from the Excel file
            ReadAndDisplayExcelFile(filePath);
        }

        static void CreateAndStyleExcelFile(string filePath)
        {
            // Create a new workbook
            using (var workbook = new XLWorkbook())
            {
                // Add a worksheet
                var worksheet = workbook.Worksheets.Add("Sheet1");

                // Add data to cells
                worksheet.Cell("A1").Value = "Name";
                worksheet.Cell("B1").Value = "Age";
                worksheet.Cell("A2").Value = "Alice";
                worksheet.Cell("B2").Value = 30;
                worksheet.Cell("A3").Value = "Bob";
                worksheet.Cell("B3").Value = 25;

                // Apply styles to headers
                var headerRange = worksheet.Range("A1:B1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.Cyan;
                headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Apply styles to cells
                worksheet.Column("A").Style.Font.FontColor = XLColor.DarkRed;
                worksheet.Column("B").Style.Fill.BackgroundColor = XLColor.LightYellow;

                // Save the workbook
                workbook.SaveAs(filePath);
            }

            Console.WriteLine($"Excel file '{filePath}' created successfully!");
        }

        static void ReadAndDisplayExcelFile(string filePath)
        {
            // Open the existing workbook
            using (var workbook = new XLWorkbook(filePath))
            {
                // Get the first worksheet
                var worksheet = workbook.Worksheet(1);

                // Read the data
                Console.WriteLine("Reading data from the Excel file:");
                for (int row = 1; row <= 3; row++)
                {
                    for (int col = 1; col <= 2; col++)
                    {
                        var cellValue = worksheet.Cell(row, col).Value;
                        Console.Write($"{cellValue} \t");
                    }
                    Console.WriteLine();
                }
            }
        }
    }
}
