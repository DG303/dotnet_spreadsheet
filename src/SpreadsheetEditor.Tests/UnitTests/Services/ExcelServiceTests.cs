using System;
using System.IO;
using Xunit;
using Moq;
using OfficeOpenXml;
using SpreadsheetEditor.Core.Services;
using SpreadsheetEditor.Core.Interfaces;

namespace SpreadsheetEditor.Tests.UnitTests.Services
{
    public class ExcelServiceTests : IDisposable
    {
        private readonly string _testFilePath;
        private readonly IExcelService _excelService;

        public ExcelServiceTests()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _testFilePath = Path.Combine(Path.GetTempPath(), $"test_excel_{Guid.NewGuid()}.xlsx");
            _excelService = new ExcelService();
        }

        public void Dispose()
        {
            if (File.Exists(_testFilePath))
            {
                File.Delete(_testFilePath);
            }
        }

        [Fact]
        public void CreateWorkbook_ShouldCreateNewExcelFile()
        {
            // Act
            _excelService.CreateWorkbook(_testFilePath);

            // Assert
            Assert.True(File.Exists(_testFilePath));
        }

        [Fact]
        public void WriteCell_ShouldUpdateCellValue()
        {
            // Arrange
            const string sheetName = "Sheet1";
            const string cellAddress = "A1";
            const string testValue = "Test Value";

            // Act
            _excelService.CreateWorkbook(_testFilePath);
            _excelService.WriteCell(_testFilePath, sheetName, cellAddress, testValue);

            // Assert
            var actualValue = _excelService.ReadCell(_testFilePath, sheetName, cellAddress);
            Assert.Equal(testValue, actualValue);
        }

        [Fact]
        public void ReadCell_WithNonExistentFile_ShouldThrowFileNotFoundException()
        {
            // Arrange
            const string nonExistentFile = "nonexistent.xlsx";

            // Act & Assert
            Assert.Throws<FileNotFoundException>(() => 
                _excelService.ReadCell(nonExistentFile, "Sheet1", "A1"));
        }

        [Fact]
        public void WriteCell_WithNonExistentSheet_ShouldCreateNewSheet()
        {
            // Arrange
            const string newSheetName = "NewSheet";
            const string cellAddress = "A1";
            const string testValue = "Test Value";

            // Act
            _excelService.CreateWorkbook(_testFilePath);
            _excelService.WriteCell(_testFilePath, newSheetName, cellAddress, testValue);

            // Assert
            var actualValue = _excelService.ReadCell(_testFilePath, newSheetName, cellAddress);
            Assert.Equal(testValue, actualValue);
        }
    }
} 