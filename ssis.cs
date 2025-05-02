using System.Data;
using Microsoft.SqlServer.Dts.Pipeline.Wrapper;
using Microsoft.SqlServer.Dts.Runtime.Wrapper;
// Add these using statements for OpenXML
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;
using System; // Needed for Convert
 
 
/// <summary>
/// Retrieves the value of a cell as a string, handling different cell types.
/// </summary>
/// <param name="document">The SpreadsheetDocument object.</param>
/// <param name="cell">The Cell object.</param>
/// <returns>The string representation of the cell value.</returns>
private string GetCellValueAsString(SpreadsheetDocument document, Cell cell)
{
    if (cell == null || cell.CellValue == null)
    {
        return null; // Or return string.Empty if you prefer "" for nulls
    }
 
    string value = cell.CellValue.InnerText;
 
    // If the cell represents a shared string, look it up
    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
    {
        SharedStringTablePart stringTablePart = document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
        if (stringTablePart != null && stringTablePart.SharedStringTable != null)
        {
            // Check if the index is valid
            if (int.TryParse(value, out int sharedStringIndex))
            {
                SharedStringItem item = stringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAtOrDefault(sharedStringIndex);
                if (item != null && item.Text != null)
                {
                    return item.Text.Text;
                }
                else if (item != null && item.InnerXml != null) // Handle rich text formatting if present
                {
                    return item.InnerXml; // May include XML tags, adjust if needed
                }
            }
        }
        // If lookup fails for some reason, return the index itself or null/empty
         return null; // Or string.Empty or the index value
    }
    // If the cell is an inline string
    else if (cell.DataType != null && cell.DataType.Value == CellValues.InlineString)
    {
         // InlineString contains Text element(s) within an <is> tag
         Text text = cell.Descendants<Text>().FirstOrDefault();
         return (text != null) ? text.Text : cell.InnerText; // Fallback to InnerText if Text tag is missing
    }
     // For numbers, dates (stored as numbers), booleans ("0" or "1"), errors, or general format
    else
    {
        // Simply return the InnerText which holds the raw value as stored in XML
        // This might be a number, a date serial number, TRUE/FALSE, 0/1 for boolean etc.
        // It might be in scientific notation for large numbers.
        return value;
    }
}
 
public override void CreateNewOutputRows()
{
    // Get variable values
    string filePath = this.Variables.ExcelFilePath; // Use the variable name you defined
    string sheetName = this.Variables.ExcelSheetName; // Optional: Use if you have a sheet name variable
 
    // Basic validation
    if (string.IsNullOrEmpty(filePath) || !System.IO.File.Exists(filePath))
    {
        bool fireAgain = false;
        this.ComponentMetaData.FireError(0, "Script Component Source", $"Excel file path is invalid or file not found: {filePath}", "", 0, out fireAgain);
        return; // Stop processing
    }
 
    try
    {
        // Open the Excel file using Open XML SDK (read-only recommended)
        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filePath, false)) // false = read-only
        {
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            WorksheetPart worksheetPart = null;
 
            // Find the sheet by name
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == sheetName);
 
            if (sheet == null)
            {
                // Optional: Fallback to the first sheet if name not found or not provided
                sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                if(sheet == null)
                {
                     bool fireAgain = false;
                     this.ComponentMetaData.FireError(0, "Script Component Source", $"Sheet '{sheetName}' not found and no sheets exist in the workbook.", "", 0, out fireAgain);
                     return;
                }
                // Log warning if falling back?
                // ComponentMetaData.FireWarning(...);
            }
 
            worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
 
 
            if (worksheetPart != null && worksheetPart.Worksheet != null)
            {
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                bool isHeaderRow = true; // Flag to skip the first row
 
                foreach (Row row in sheetData.Elements<Row>())
                {
                    // --- Optional: Skip Header Row ---
                    if (isHeaderRow)
                    {
                        isHeaderRow = false; // Set flag to false after skipping the first row
                        continue; // Skip processing this row
                    }
                    // --- End Optional: Skip Header Row ---
 
 
                    // Add a new row to the output buffer
                    Output0Buffer.AddRow();
 
                    // Assuming columns are sequential A, B, C...
                    // You might need more robust logic if columns are sparse
                    var cells = row.Elements<Cell>().ToList();
                    int cellCounter = 0; // Use if accessing buffer columns by index
 
                    // --- Assign values to output buffer columns ---
                    // IMPORTANT: Match Buffer.<YourOutputColumnName> with the names defined in SSIS Output Columns
                    //            Access cells carefully, assuming order or using CellReference
 
                    // Example: Reading first three columns (A, B, C)
                    // Adjust CellReference or index logic if needed
 
                    Cell cellA = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value != null && c.CellReference.Value.StartsWith("A"));
                    Output0Buffer.ColumnA = GetCellValueAsString(spreadsheetDocument, cellA) ?? string.Empty; // Use ?? string.Empty to avoid nulls if required by downstream components
 
                    Cell cellB = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value != null && c.CellReference.Value.StartsWith("B"));
                    Output0Buffer.ColumnB = GetCellValueAsString(spreadsheetDocument, cellB) ?? string.Empty;
 
                    Cell cellC = row.Elements<Cell>().FirstOrDefault(c => c.CellReference?.Value != null && c.CellReference.Value.StartsWith("C"));
                    Output0Buffer.ColumnC = GetCellValueAsString(spreadsheetDocument, cellC) ?? string.Empty;
 
                    // Add more lines like the above for all your expected columns (ColumnD, ColumnE, etc.)
                    // Ensure Output0Buffer.YourColumnName matches exactly what you defined.
                }
            }
             else
            {
                 bool fireAgain = false;
                 this.ComponentMetaData.FireError(0, "Script Component Source", "Could not find worksheet data.", "", 0, out fireAgain);
             }
        } // using SpreadsheetDocument ensures it's closed and disposed
    }
    catch (System.IO.IOException ioEx)
    {
         // Handle file locking issues
         bool fireAgain = false;
         this.ComponentMetaData.FireError(0, "Script Component Source", $"Error accessing Excel file (might be open?): {ioEx.Message}", "", 0, out fireAgain);
    }
    catch (Exception ex)
    {
        // General error handling
        bool fireAgain = false;
        this.ComponentMetaData.FireError(0, "Script Component Source", $"An error occurred: {ex.Message}\nStackTrace: {ex.StackTrace}", "", 0, out fireAgain);
    }
}
 
