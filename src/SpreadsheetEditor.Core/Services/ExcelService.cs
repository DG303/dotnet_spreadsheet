using System;
using System.IO;
using OfficeOpenXml;
using SpreadsheetEditor.Core.Interfaces;
using SpreadsheetEditor.Core.Models;

namespace SpreadsheetEditor.Core.Services
{
    public class ExcelService : IExcelService
    {
        static ExcelService()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public async Task CreateWorkbookAsync(string filePath, string? sheetName = null)
        {
            using var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add(sheetName ?? "Sheet1");
            await package.SaveAsAsync(filePath);
        }

        public async Task WriteCellAsync(string filePath, string cellAddress, object value, string? sheetName = null)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Excel file not found: {filePath}");
            }

            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[sheetName ?? "Sheet1"];
            
            worksheet.Cells[cellAddress].Value = value;
            await package.SaveAsync();
        }

        public async Task<object?> ReadCellAsync(string filePath, string cellAddress, string? sheetName = null)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Excel file not found: {filePath}");
            }

            using var package = new ExcelPackage(new FileInfo(filePath));
            var worksheet = package.Workbook.Worksheets[sheetName ?? "Sheet1"];
            
            if (worksheet == null)
            {
                throw new ArgumentException($"Worksheet '{sheetName}' not found in the workbook");
            }

            return worksheet.Cells[cellAddress].Value;
        }
    }
} 