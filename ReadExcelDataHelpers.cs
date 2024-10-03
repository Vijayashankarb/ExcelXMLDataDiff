using System;
using System.IO;
using OfficeOpenXml;
namespace ExcelToXML;
internal class ReadExcelDataHelpers
{
    public void ReadExcelData()
    {
    // Path to your Excel file
    //C: \Users\vsha\source\repos\ExcelToXML\TestData\
        string filePath = Path.Combine(Environment.CurrentDirectory, @"TestData\", "SampleTestData.xlsx");

        // Check if the file exists
        if (!File.Exists(filePath))
        {
            Console.WriteLine("Excel File not found.");
            return;
        }

        // If you use EPPlus in a noncommercial context
        // according to the Polyform Noncommercial license:
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


        // Load the Excel file
        FileInfo fileInfo = new FileInfo(filePath);
        using (ExcelPackage package = new ExcelPackage(fileInfo))
        {
            // Get the first worksheet
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

            // Get the number of rows and columns
            int rowCount = worksheet.Dimension.Rows;
            int colCount = worksheet.Dimension.Columns;

            // Read and display the contents
            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    // Read the cell value
                    var cellValue = worksheet.Cells[row, col].Value;
                    Console.Write($"{cellValue}\t");
                }
                Console.WriteLine();
            }
        }
    }

    public static IEnumerable<string> GetColumnValues(string fileName, string columnLetter)
    {

        string filePath = Path.Combine(Environment.CurrentDirectory, @"TestData\", fileName);

        // Check if the file exists
        if (!File.Exists(filePath))
        {
            Console.WriteLine("Excel File not found.");
            return null;
        }

        // If you use EPPlus in a noncommercial context
        // according to the Polyform Noncommercial license:
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


        // Load the Excel file
        FileInfo fileInfo = new FileInfo(filePath);

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.End.Row;

            // Use LINQ to select column values
            return Enumerable.Range(1, rowCount)
                             .Select(row => worksheet.Cells[$"{columnLetter}{row}"].Text)
                             .Where(value => !string.IsNullOrEmpty(value))
                             .ToList(); // Convert to list or return as IEnumerable
        }
    }



}