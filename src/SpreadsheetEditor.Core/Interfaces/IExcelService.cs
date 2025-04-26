using SpreadsheetEditor.Core.Models;

namespace SpreadsheetEditor.Core.Interfaces
{
    public interface IExcelService
    {
        /// <summary>
        /// Creates a new Excel workbook at the specified path
        /// </summary>
        /// <param name="filePath">The path where the workbook should be created</param>
        void CreateWorkbook(string filePath);

        /// <summary>
        /// Writes a value to a specific cell in an Excel workbook
        /// </summary>
        /// <param name="filePath">Path to the Excel workbook</param>
        /// <param name="sheetName">Name of the worksheet</param>
        /// <param name="cellAddress">Address of the cell (e.g., "A1")</param>
        /// <param name="value">Value to write to the cell</param>
        void WriteCell(string filePath, string sheetName, string cellAddress, object value);

        /// <summary>
        /// Reads the value from a specific cell in an Excel workbook
        /// </summary>
        /// <param name="filePath">Path to the Excel workbook</param>
        /// <param name="sheetName">Name of the worksheet</param>
        /// <param name="cellAddress">Address of the cell (e.g., "A1")</param>
        /// <returns>The value of the specified cell</returns>
        object ReadCell(string filePath, string sheetName, string cellAddress);

        Task CreateWorkbookAsync(string filePath, string? sheetName = null);
        Task WriteCellAsync(string filePath, string cellAddress, object value, string? sheetName = null);
        Task<object?> ReadCellAsync(string filePath, string cellAddress, string? sheetName = null);
    }
} 