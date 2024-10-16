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
        private const string DefaultRangeFile = $"{DefaultDirectory}/spreadsheet_range.xlsx";
        private const string DefaultMergeCellsFilePath = $"{DefaultDirectory}/spreadsheet_merged_cells.xlsx";
        private const string DefaultDataPopulatedFilePath = $"{DefaultDirectory}/spreadsheet_data_populated.xlsx";
        private const string DefaultProtectedSheetFile = $"{DefaultDirectory}/spreadsheet_protected_sheet.xlsx";
        private const string DefaultUnProtectedSheetFile = $"{DefaultDirectory}/spreadsheet_un_protected_sheet.xlsx";
        private const string DefaultFileDemonstrateBasics = $"{DefaultDirectory}/spreadsheet_worksheet_demonstrate_basics.xlsx";
        private const string DefaultFileWithImage = $"{DefaultDirectory}/spreadsheet_images.xlsx";
        private const string DefaultImageDirectory = "../../../spreadSheetImages";
        private const string DefaultImageFile = $"{DefaultImageDirectory}/image1.png";
        private const string DefaultSpreadSheetFileData = "../../../defaultSpreadSheets/defaultSpreadSheet.xlsx";
        private const string BasicOperationFile = $"{DefaultDirectory}/spreadsheet_basic_operations.xlsx";
        private const string BasicOperationFileColumns = $"{DefaultDirectory}/spreadsheet_basic_operations_columns.xlsx";
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

        public void AddImageToWorksheet(string filePath = DefaultFileWithImage, string imagePath = DefaultImageFile)
        {
            try
            {
                using (var workbook = new Workbook())
                {
                    // Adding an image to the first worksheet
                    var firstSheet = workbook.Worksheets[0];
                    // Assuming the Image class and AddImage method exist and work as described
                    var imageForFirstSheet = new Image(imagePath); // Create an instance of Image class
                    firstSheet.AddImage(imageForFirstSheet, 6, 1, 8, 3); // Add image to worksheet



                    // Saving the workbook with the added images
                    workbook.Save(filePath);
                    Console.WriteLine($"Workbook saved with images at {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while adding images to the workbook: {ex.Message}");
            }
        }

        public void ExtractImagesFromWorksheet(string filePath = DefaultFileWithImage, string outputDirectory = DefaultImageDirectory)
        {
            try
            {
                Workbook wb = new Workbook(filePath); // Load the workbook

                // Select the first worksheet
                var worksheet = wb.Worksheets[0];

                // Extract images
                var images = worksheet.ExtractImages(); // Assuming ExtractImages method exists and returns image objects

                // Ensure the output directory exists
                if (!Directory.Exists(outputDirectory))
                {
                    Directory.CreateDirectory(outputDirectory);
                    Console.WriteLine($"Created directory at {outputDirectory}");
                }

                // Save each extracted image
                foreach (var image in images)
                {
                    var outputFilePath = Path.Combine(outputDirectory, $"Image_{Guid.NewGuid()}.{image.Extension}");
                    using (var fileStream = File.Create(outputFilePath))
                    {
                        image.Data.CopyTo(fileStream); // Assuming image object has Data (Stream) and Extension properties
                    }

                    Console.WriteLine($"Image saved to {outputFilePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public void SetColumnWidthAndRowHeight(string filePath = DefaultFileDemonstrateBasics)
        {
            // Initialize a new workbook instance
            using (Workbook wb = new Workbook())
            {
                // Access the first worksheet in the workbook
                Worksheet firstSheet = wb.Worksheets[0];

                // Set the height of the first row to 40 points
                firstSheet.SetRowHeight(1, 40);

                // Set the width of column B to 75 points
                firstSheet.SetColumnWidth("B", 75);

                // Insert a value into cell A1
                Cell cellA1 = firstSheet.Cells["A1"];
                cellA1.PutValue("Value in A1");

                Cell cellB2 = firstSheet.Cells["B2"];
                cellB2.PutValue("Text in B2");

                // Save the workbook to the specified file path
                wb.Save(filePath);

                Console.WriteLine($"Workbook saved successfully at {filePath}");
            }
        }

        public void SetRangeValue(string filePath = DefaultRangeFile)
        {
           
            using (Workbook wb = new Workbook())
            {
                // Access the first worksheet in the workbook
                Worksheet firstSheet = wb.Worksheets[0];

                // Select a range within the worksheet
                var range = firstSheet.GetRange("A1", "B10");
                Console.WriteLine($"Column count: {range.ColumnCount}");
                Console.WriteLine($"Row count: {range.RowCount}");

                // Set a similar value to all cells in the selected range
                range.SetValue("Hello");

                // Save the changes back to the workbook
                wb.Save(filePath);

                Console.WriteLine("Value set to range and workbook saved successfully.");
            }
        }

        public void InsertRowsIntoWorksheet(string filePath = DefaultSpreadSheetFileData)
        {
           
            // Load the workbook from the specified file path
            using (Workbook wb = new Workbook(filePath))
            {
                var saveFilePath = BasicOperationFile;
                // Access the first worksheet in the workbook
                Worksheet firstSheet = wb.Worksheets[0];

                // Define the starting row index and the number of rows to insert
                uint startRowIndex = 5;
                uint numberOfRows = 3;

                // Insert the rows into the worksheet
                firstSheet.InsertRows(startRowIndex, numberOfRows);

                // Get the total row count after insertion
                int rowsCount = firstSheet.GetRowCount();

                // Output the updated row count to the console
                Console.WriteLine("Rows Count=" + rowsCount);

                // Save the workbook to reflect the changes made
                wb.Save(saveFilePath);

                Console.WriteLine("Rows inserted and workbook saved successfully.");
            }
        }

        public void InsertColumnsIntoWorksheet(string filePath = DefaultSpreadSheetFileData)
        {
            // Load the workbook from the specified file path
            using (Workbook wb = new Workbook(filePath))
            {
                var saveFilePath = BasicOperationFileColumns;

                // Access the first worksheet in the workbook
                Worksheet firstSheet = wb.Worksheets[0];

                // Define the starting column and the number of columns to insert
                string startColumn = "B";
                int numberOfColumns = 3;

                // Insert the columns into the worksheet
                firstSheet.InsertColumns(startColumn, numberOfColumns);

                // Get the total column count after insertion (Note: Moved after insertion for correct count)
                int columnsCount = firstSheet.GetColumnCount();

                // Output the updated column count to the console
                Console.WriteLine("Columns Count=" + columnsCount);

                // Save the workbook to reflect the changes made
                wb.Save(saveFilePath);

                Console.WriteLine("Columns inserted and workbook saved successfully.");
            }
        }

        public void GetHiddenColumns(string filePath = DefaultSpreadSheetFileData)
        {
            // Load the workbook from the specified file path
            using (Workbook wb = new Workbook(filePath))
            {
                var saveFilePath = BasicOperationFileColumns;

                // Access the first worksheet in the workbook
                Worksheet firstSheet = wb.Worksheets[0];

                // Get Hidden Columns
                List<uint> hiddenColumns = firstSheet.GetHiddenColumns();

                foreach (var col in hiddenColumns)
                {
                    Console.WriteLine($"Hidden Column: {col}");
                }

            }
        }

        public void GetHiddenRows(string filePath = DefaultSpreadSheetFileData)
        {
            // Load the workbook from the specified file path
            using (Workbook wb = new Workbook(filePath))
            {
                var saveFilePath = BasicOperationFileColumns;

                // Access the first worksheet in the workbook
                Worksheet firstSheet = wb.Worksheets[0];

                // Get Hidden Columns
                List<uint> hiddenRows = firstSheet.GetHiddenColumns();

                foreach (var row in hiddenRows)
                {
                    Console.WriteLine($"Hidden Row: {row}");
                }

            }
        }

        public void FreezePane(string filePath = DefaultSpreadSheetFileData)
        {
            // Load the workbook from the specified file path
            using (Workbook wb = new Workbook(filePath))
            {
                var saveFilePath = BasicOperationFileColumns;

                // Access the first worksheet in the workbook
                Worksheet firstSheet = wb.Worksheets[0];

                firstSheet.FreezePane(2,1);

            }
        }
    }
}
