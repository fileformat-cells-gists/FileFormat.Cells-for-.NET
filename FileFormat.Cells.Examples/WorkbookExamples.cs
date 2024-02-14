namespace FileFormat.Cells.Examples
{
    public class WorkbookExamples
    {
        private const string DefaultDirectory = "../../../spreadSheetDocuments/Workbook";
        private const string DefaultFilePath = $"{DefaultDirectory}/spreadsheet.xlsx"; 
        private const string DefaultStyledFilePath = $"{DefaultDirectory}/spreadsheet_default_styled.xlsx";
        private const string MultipleStyledFilePath = $"{DefaultDirectory}/spreadsheet_multiple_styled.xlsx";
        private const string DefaultFilePathForProperties = $"{DefaultDirectory}/spreadsheet_properties.xlsx";
        private const string DefaultFileWithImage = $"{DefaultDirectory}/spreadsheet_images.xlsx";
        private const string DefaultImageDirectory = "../../../spreadSheetImages";
        private const string DefaultImageFile = $"{DefaultImageDirectory}/image1.png";
        private const string DefaultSheetName = "NewWorksheet";

        public WorkbookExamples()
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
        public void CreateAndSaveWorkbook(string filePath = DefaultFilePath)
        {
            try
            {
                // Initialize an instance of the Aspose.Cells Workbook class
                Workbook workbook = new Workbook();

                // Here you can add data to the workbook, format cells, etc.

                // Save the workbook to the specified path
                workbook.Save(filePath);
            }
            catch (Exception ex)
            {
                // Handle the exception, log it or print to console
                Console.WriteLine($"An error occurred while creating or saving the workbook: {ex.Message}");
            }
        }

        public void AddSheetToWorkbook(string filePath = DefaultFilePath, string sheetName = DefaultSheetName)
        {
            try
            {
                // Open the existing workbook.
                using (var workbook = new FileFormat.Cells.Workbook(filePath))
                {
                    // Add a new worksheet to the workbook with the provided name.
                    Worksheet newSheet = workbook.AddSheet(sheetName);

                    // Example content to the new worksheet.
                    Cell cellA1 = newSheet.Cells["A1"];
                    cellA1.PutValue("Hello from the new sheet!");

                    // Save the workbook with the added worksheet.
                    workbook.Save(filePath);
                }
            }
            catch (Exception ex)
            {
                // Handle the exception, log it or print to console
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        public void CreateStyledWorkbook(string filePath = MultipleStyledFilePath)
        {
            try
            {
                // Creating a new workbook instance
                using (Workbook wb = new Workbook())
                {
                    // Create a custom style with Arial font, size 11, and red color
                    uint styleIndex = wb.CreateStyle("Arial", 11, "FF0000");

                    // Create another custom style with Times New Roman font, size 12, and black color
                    uint styleIndex2 = wb.CreateStyle("Times New Roman", 12, "000000");

                    // Access the first worksheet from the workbook
                    Worksheet firstSheet = wb.Worksheets[0];

                    // Assign a value to the cell A1 and apply the first custom style
                    Cell cellA1 = firstSheet.Cells["A1"];
                    cellA1.PutValue("Styled Text A1");
                    cellA1.ApplyStyle(styleIndex);

                    // Assign a value to the cell B2 and apply the second custom style
                    Cell cellB2 = firstSheet.Cells["B2"];
                    cellB2.PutValue("Styled Text B2");
                    cellB2.ApplyStyle(styleIndex2);

                    // Save the workbook to the specified file path
                    wb.Save(filePath);
                }
            }
            catch (Exception ex)
            {
                // Handle the exception, log it or print to console
                Console.WriteLine($"An error occurred while creating or saving the styled workbook: {ex.Message}");
            }
        }

        public void CreateWorkbookWithDefaultStyle(string filePath = DefaultStyledFilePath)
        {
            try
            {
                // Open the existing workbook specified by the filePath
                using (Workbook wb = new Workbook())
                {
                    // Update the default style for the workbook
                    wb.UpdateDefaultStyle("Arial", 12, "A02000"); // Set font to Arial, size 12, and a shade of red

                    // Get the first worksheet from the workbook
                    Worksheet firstSheet = wb.Worksheets[0];

                    // Add values to the cells A1 to A3 with the new default style
                    firstSheet.Cells["A1"].PutValue("Hello, World!");
                    firstSheet.Cells["A2"].PutValue("Default Style Applied!");
                    firstSheet.Cells["A3"].PutValue("Check Font & Color.");

                    // Save the changes made to the new workbook
                    wb.Save(filePath);
                }
            }
            catch (Exception ex)
            {
                // Handle the exception, log it or print to console
                Console.WriteLine($"An error occurred while creating or saving the workbook with default style: {ex.Message}");
            }
        }

        public void CreateWorkbookWithProperties(string filePath = DefaultFilePathForProperties)
        {
            try
            {
                // Create a new workbook and set cell values and document properties
                using (var workbook = new Workbook())
                {
                    // Access the first worksheet
                    Worksheet firstSheet = workbook.Worksheets[0];

                    // Set values for cells A1 and A2
                    firstSheet.Cells["A1"].PutValue("Text A1");
                    firstSheet.Cells["A2"].PutValue("Text A2");

                    // Configure document properties
                    var newProperties = new BuiltInDocumentProperties
                    {
                        Author = "Fahad Adeel",
                        Title = "Sample Workbook",
                        CreatedDate = DateTime.Now,
                        ModifiedBy = "Fahad",
                        ModifiedDate = DateTime.Now.AddHours(1),
                        Subject = "Testing Subject"
                    };

                    // Assign the new properties to the workbook
                    workbook.BuiltinDocumentProperties = newProperties;

                    // Save the workbook to the specified path
                    workbook.Save(filePath);
                }
            }
            catch (Exception ex)
            {
                // Handle the exception, log it or print to console
                Console.WriteLine($"An error occurred while creating or saving the workbook with properties: {ex.Message}");
            }
        }

        public void DisplayWorkbookProperties(string filePath = DefaultFilePathForProperties)
        {
            try
            {
                // Open the existing workbook
                using (var workbook = new Workbook(filePath))
                {
                    // Retrieve and display the document properties
                    var properties = workbook.BuiltinDocumentProperties;
                    Console.WriteLine($"Author: {properties.Author}");
                    Console.WriteLine($"Title: {properties.Title}");
                    Console.WriteLine($"Created Date: {properties.CreatedDate}");
                    Console.WriteLine($"Modified By: {properties.ModifiedBy}");
                    Console.WriteLine($"Modified Date: {properties.ModifiedDate}");
                    Console.WriteLine($"Subject: {properties.Subject}");
                    Console.WriteLine("=================================");
                }
            }
            catch (Exception ex)
            {
                // Handle the exception, log it or print to console
                Console.WriteLine($"An error occurred while opening the workbook: {ex.Message}");
            }
        }

        public void RemoveWorksheetByName(string filePath = DefaultFilePath, string sheetName = "sheet1")
        {
            try
            {
                using (var workbook = new Workbook(filePath))
                {
                    // Attempt to remove the specified worksheet
                    bool removed = workbook.RemoveSheet(sheetName);

                    if (removed)
                    {
                        // Save the changes to the workbook
                        workbook.Save();
                        Console.WriteLine($"{sheetName} removed and changes saved successfully!");
                    }
                    else
                    {
                        Console.WriteLine($"The specified worksheet {sheetName} was not found.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while removing the worksheet: {ex.Message}");
            }
        }


    }
}
