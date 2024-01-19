using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text;

namespace FileFormat.Cells.Examples
{
    public class WorksheetExamples
    {
        private const string DefaultDirectory = "../../../spreadSheetDocuments/Worksheet";
        private const string DefaultFilePath = $"{DefaultDirectory}/spreadsheet.xlsx";
        private const string DefaultMergeCellsFilePath = $"{DefaultDirectory}/spreadsheet_merged_cells.xlsx";
        private const string DefaultDataPopulatedFilePath = $"{DefaultDirectory}/spreadsheet_data_populated.xlsx";
        private const string DefaultProtectedSheetFile = $"{DefaultDirectory}/spreadsheet_protected_sheet.xlsx";
        private const string DefaultUnProtectedSheetFile = $"{DefaultDirectory}/spreadsheet_un_protected_sheet.xlsx";
        
        public WorksheetExamples()
        {
            if (!System.IO.Directory.Exists(DefaultDirectory))
            {
                // If it doesn't exist, create the directory
                System.IO.Directory.CreateDirectory(DefaultDirectory);
                System.Console.WriteLine($"Directory '{System.IO.Path.GetFullPath(DefaultDirectory)}' " +
                    $"created successfully.");
            }
            else
            {
                var files = System.IO.Directory.GetFiles(System.IO.Path.GetFullPath(DefaultDirectory));
                foreach (var file in files)
                {
                    System.IO.File.Delete(file);
                    System.Console.WriteLine($"File deleted: {file}");
                }
                System.Console.WriteLine($"Directory '{System.IO.Path.GetFullPath(DefaultDirectory)}' " +
                    $"cleaned up.");
            }
        }

        public void AddValueToCell(string filePath = DefaultFilePath, string cellName = "A1", string value = "Text in A1")
        {
            try
            {
                using (Workbook wb = new Workbook())
                {
                    // Accessing the first worksheet
                    Worksheet firstSheet = wb.Worksheets.First();

                    // Adding value to the specified cell
                    Cell cell = firstSheet.Cells[cellName];
                    cell.PutValue(value);

                    // Save the changes to the workbook
                    wb.Save(filePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public void ReadCellValue(string filePath = DefaultFilePath, string cellName = "A1")
        {
            try
            {
                using (Workbook wb = new Workbook(filePath)) // Load existing workbook
                {
                    Worksheet firstSheet = wb.Worksheets[0]; // Access the first worksheet
                    Cell cell = firstSheet.Cells[cellName]; // Get the specified cell

                    // Output cell data type and value
                    Console.WriteLine($"Data Type of {cellName}: {cell.GetDataType()}");
                    string value = cell.GetValue(); // Get the value in the cell
                    Console.WriteLine($"Value in {cellName}: {value}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public void AddFormulasToCells(string filePath = DefaultFilePath)
        {
            try
            {
                using (Workbook wb = new Workbook())
                {
                    Worksheet firstSheet = wb.Worksheets[0];
                    Random rand = new Random();

                    Console.WriteLine("Adding random values to cells B1 to B10...");

                    for (int i = 1; i <= 10; i++)
                    {
                        string cellReference = $"B{i}";
                        Cell cell = firstSheet.Cells[cellReference];
                        double randomValue = rand.Next(1, 100);
                        cell.PutValue(randomValue);

                        Console.WriteLine($"Cell {cellReference} set to {randomValue}");
                    }

                    Console.WriteLine("Adding a formula to cell B11 to sum values from B1 to B10...");
                    Cell cellA11 = firstSheet.Cells["B11"];
                    cellA11.PutFormula("SUM(B1:B10)");

                    wb.Save(filePath);
                    Console.WriteLine($"Workbook saved with formulas at {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public void MergeCellsInWorksheet(string filePath = DefaultMergeCellsFilePath)
        {
            try
            {
                using (var workbook = new Workbook())
                {
                    var firstSheet = workbook.Worksheets[0];
                    Console.WriteLine("Merging cells A1 to C1...");

                    firstSheet.MergeCells("A1", "C1"); // Merge cells from A1 to C1

                    // Add value to the top-left cell of the merged area
                    var topLeftCell = firstSheet.Cells["A1"];
                    topLeftCell.PutValue("This is a merged cell");

                    workbook.Save(filePath);
                    Console.WriteLine($"Workbook saved with merged cells at {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public void CreateStyledWorkbookWithSampleData(string filePath = DefaultDataPopulatedFilePath)
        {
            try
            {
                using (var workbook = new Workbook())
                {
                    // Set the default style for the workbook
                    workbook.UpdateDefaultStyle("Times New Roman", 11, "000000");

                    // Define styles for headers, even rows, and odd rows
                    uint headerStyleIndex = workbook.CreateStyle("Arial Black", 15, "4F81BD"); // Dark Blue text for headers
                    uint evenRowStyleIndex = workbook.CreateStyle("Arial", 12, "FF0000"); // Red text for even rows
                    uint oddRowStyleIndex = workbook.CreateStyle("Calibri", 12, "0000FF"); // Blue text for odd rows

                    // Get the first worksheet in the workbook
                    var firstSheet = workbook.Worksheets[0];

                    // Define header labels and apply the header style
                    string[] headers = { "Student ID", "Student Name", "Course", "Grade" };
                    for (int col = 0; col < headers.Length; col++)
                    {
                        string cellAddress = $"{(char)(65 + col)}1";
                        Cell cell = firstSheet.Cells[cellAddress];
                        cell.PutValue(headers[col]);
                        cell.ApplyStyle(headerStyleIndex);
                    }

                    // Populate the worksheet with sample data and apply styles for even and odd rows
                    int rowCount = 10;
                    for (int row = 2; row <= rowCount + 1; row++)
                    {
                        for (int col = 0; col < headers.Length; col++)
                        {
                            string cellAddress = $"{(char)(65 + col)}{row}";
                            Cell cell = firstSheet.Cells[cellAddress];

                            // Generate sample data based on the column index
                            switch (col)
                            {
                                case 0: // Student ID
                                    cell.PutValue($"ID{row - 1}");
                                    break;
                                case 1: // Student Name
                                    cell.PutValue($"Student {row - 1}");
                                    break;
                                case 2: // Course
                                    cell.PutValue($"Course {(row - 1) % 5 + 1}");
                                    break;
                                case 3: // Grade
                                    cell.PutValue($"Grade {((row - 1) % 3) + 'A'}");
                                    break;
                            }

                            // Apply style based on whether the row number is even or odd
                            cell.ApplyStyle((row % 2 == 0) ? evenRowStyleIndex : oddRowStyleIndex);
                        }
                    }

                    // Save the changes to the workbook
                    workbook.Save(filePath);

                    Console.WriteLine($"Workbook saved with sample data at {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public void ProtectWorksheet(string filePath = DefaultProtectedSheetFile, string password = "123456")
        {
            try
            {
                using (var workbook = new Workbook())
                {
                    Worksheet sheetToProtect = workbook.Worksheets[0];

                    Console.WriteLine("Protecting the first worksheet...");

                    sheetToProtect.ProtectSheet(password);

                    workbook.Save(filePath);
                    Console.WriteLine($"Worksheet protected and workbook saved at {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public void UnprotectWorksheets(string protectedFilePath = DefaultProtectedSheetFile, string filePath = DefaultUnProtectedSheetFile)
        {
            try
            {
                using (Workbook wb = new Workbook(protectedFilePath))
                {
                    Console.WriteLine("Checking for protected worksheets...");

                    foreach (var worksheet in wb.Worksheets)
                    {
                        if (worksheet.IsProtected())
                        {
                            Console.WriteLine("Protected Sheet Name = " + worksheet.Name);
                            worksheet.UnprotectSheet();

                            Console.WriteLine($"Unprotected worksheet: {worksheet.Name}");
                        }
                    }

                    wb.Save(filePath);
                    Console.WriteLine($"All protected worksheets are now unprotected. Workbook saved at {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
