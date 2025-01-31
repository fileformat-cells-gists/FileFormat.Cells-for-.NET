﻿
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

            // Call the method to Remove sheet from a workbook.
            myWorkbook.RemoveWorksheetByName();

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

            // Call the method to add image within a worksheet.
            myWorksheet.AddImageToWorksheet();

            // Call the method to extract images within a worksheet.
            myWorksheet.ExtractImagesFromWorksheet();

            // Call the method to Set Column width and Row height within the worksheet.
            myWorksheet.SetColumnWidthAndRowHeight();

            // Call the method to set the value for a range within a worksheet.
            myWorksheet.SetRangeValue();

            // Call the method to Insert Rows within a worksheet.
            myWorksheet.InsertRowsIntoWorksheet();

            // Call the method to Insert Columns within a worksheet.
            myWorksheet.InsertColumnsIntoWorksheet();

            // Call the method to Get Hidden Columns within a worksheet.
            myWorksheet.GetHiddenColumns();

            // Call the method to Get Hidden Rows within a worksheet.
            myWorksheet.GetHiddenRows();

            // Call the method to Freeze Pane within a worksheet.
            myWorksheet.FreezePane();
        }
    }
}