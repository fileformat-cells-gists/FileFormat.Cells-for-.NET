
using FileFormat.Cells.Examples;

namespace FileForamt.Cells.Examples.Usage
{
    class Program
    {
        static void Main(string[] args)
        {

            // Create an instance of the WorkbookExamples class from FileFormat.Cells.Examples
            WorkbookExamples myWorkbook = new WorkbookExamples();
            // Call the method to create and save the workbook
            myWorkbook.CreateAndSaveWorkbook();

            // Call the method to add a new sheet to the workbook
            myWorkbook.AddSheetToWorkbook();

            // Call the method to create a workbook with custom styles
            myWorkbook.CreateStyledWorkbook();

            // Call the method to create a workbook with a default style
            myWorkbook.CreateWorkbookWithDefaultStyle();

            // Call the method to create a workbook with properties
            myWorkbook.CreateWorkbookWithProperties();

            // Call the method to display the workbook properties
            myWorkbook.DisplayWorkbookProperties();

            // Create an instance of the WorksheetExamples class from FileFormat.Cells.Examples
            WorksheetExamples myWorksheet = new WorksheetExamples();
            
            // Call the method to add a value to a cell
            myWorksheet.AddValueToCell();

            // Call the method to read and display the value of a cell
            myWorksheet.ReadCellValue();

            // Call the method to add formulas to cells
            myWorksheet.AddFormulasToCells();

            // Call the method to merge cells in a worksheet
            myWorksheet.MergeCellsInWorksheet();

            // Call the method to fill in the sample data in a worksheet.
            myWorksheet.CreateStyledWorkbookWithSampleData();

            // Call the method to protect a worksheet in the spreadsheet.
            myWorksheet.ProtectWorksheet();

            // Call the method to un protect a worksheet in the spreadsheet.
            myWorksheet.UnprotectWorksheets();
        }
    }
}