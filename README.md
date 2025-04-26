# 📄 Excel Automation with C# — Using EPPlus

This project demonstrates how to automate Excel file manipulation using **EPPlus**, a lightweight, server-safe library for `.xlsx` files in C#.

Rather than relying on Microsoft Office Interop (which requires Excel installed and is not server-friendly), **EPPlus** allows reading, writing, and formatting Excel files purely via C# code — **no Excel installation needed**.

---

## 📋 Prerequisites

- [.NET 8.0 SDK](https://dotnet.microsoft.com/download) or later
- [Visual Studio 2022](https://visualstudio.microsoft.com/) or [VS Code](https://code.visualstudio.com/) with C# extensions
- Basic understanding of C# and Excel file formats

---

## 📦 Technologies Used

- **C# / .NET 8.0**
- **EPPlus 7.0.9** (NuGet Package)
- **Visual Studio / VS Code**
- **Docker** (for containerization and testing)

---

## 📁 Project Structure

```
SpreadsheetEditor/
├── src/
│   ├── SpreadsheetEditor.Core/          # Core business logic
│   │   └── SpreadsheetEditor.Core.csproj
│   │
│   └── SpreadsheetEditor.Console/       # Console application
│       ├── Program.cs                   # Entry point
│       ├── Options.cs                   # Command-line options
│       └── SpreadsheetEditor.Console.csproj
│
├── data/                                # Excel file storage (Docker volume)
├── Dockerfile                           # Docker build configuration
├── docker-compose.yml                   # Docker Compose configuration
├── test-docker.sh                       # Docker test script
└── README.md                           # Project documentation
```

---

## 🚀 Getting Started

### 1. Create the Solution and Projects

```bash
# Create solution
dotnet new sln -n SpreadsheetEditor

# Create projects
dotnet new classlib -n SpreadsheetEditor.Core -o src/SpreadsheetEditor.Core
dotnet new console -n SpreadsheetEditor.Console -o src/SpreadsheetEditor.Console
dotnet new xunit -n SpreadsheetEditor.Tests -o src/SpreadsheetEditor.Tests

# Add projects to solution
dotnet sln add src/SpreadsheetEditor.Core/SpreadsheetEditor.Core.csproj
dotnet sln add src/SpreadsheetEditor.Console/SpreadsheetEditor.Console.csproj
dotnet sln add src/SpreadsheetEditor.Tests/SpreadsheetEditor.Tests.csproj
```

### 2. Install Required Packages

```bash
# Core project
cd src/SpreadsheetEditor.Core
dotnet add package EPPlus

# Test project
cd ../SpreadsheetEditor.Tests
dotnet add package Moq
```

### 3. Basic Implementation Example

#### Core Models (SpreadsheetEditor.Core/Models/ExcelCell.cs)
```csharp
public class ExcelCell
{
    public string Address { get; set; }
    public object Value { get; set; }
    public string Formula { get; set; }
    public CellStyle Style { get; set; }
}
```

#### Excel Service (SpreadsheetEditor.Core/Services/ExcelService.cs)
```csharp
public class ExcelService : IExcelService
{
    public void CreateWorkbook(string filePath)
    {
        using var package = new ExcelPackage();
        var worksheet = package.Workbook.Worksheets.Add("Sheet1");
        package.SaveAs(new FileInfo(filePath));
    }

    public void WriteCell(string filePath, string sheetName, string cellAddress, object value)
    {
        using var package = new ExcelPackage(new FileInfo(filePath));
        var worksheet = package.Workbook.Worksheets[sheetName];
        worksheet.Cells[cellAddress].Value = value;
        package.Save();
    }
}
```

#### Console Application (SpreadsheetEditor.Console/Program.cs)
```csharp
class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        var excelService = new ExcelService();
        excelService.CreateWorkbook("Sample.xlsx");
        excelService.WriteCell("Sample.xlsx", "Sheet1", "A1", "Hello World!");
    }
}
```

---

## 🔥 Key Features

1. **Excel File Operations**
   - Create new workbooks
   - Read and write cell values
   - Apply cell formatting
   - Handle formulas
   - Manage multiple worksheets

2. **Data Processing**
   - Import/export data
   - Data validation
   - Cell range operations
   - Conditional formatting

3. **Formatting Capabilities**
   - Font styles
   - Cell colors
   - Borders
   - Number formats
   - Alignment

---

## 🧪 Testing

The project includes an automated test suite that runs in Docker. The test script (`test-docker.sh`) performs the following operations:

1. Creates a new Excel file
2. Writes a test value to cell A1
3. Reads the value back and verifies it matches
4. Cleans up the test file

To run the tests:
```bash
chmod +x test-docker.sh
./test-docker.sh
```

---

## 📚 Resources

- [EPPlus Official Documentation](https://epplussoftware.com/docs)
- [.NET Documentation](https://docs.microsoft.com/en-us/dotnet/)
- [Excel File Format Specifications](https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk)

---

## 🧐 Best Practices

1. **Error Handling**
   - Use try-catch blocks for file operations
   - Validate file paths and permissions
   - Handle Excel-specific exceptions

2. **Performance**
   - Use `using` statements for proper resource disposal
   - Minimize file operations
   - Batch cell updates when possible

3. **Code Organization**
   - Follow SOLID principles
   - Use dependency injection
   - Implement proper logging

---

## 🌟 Next Steps

1. **Advanced Features**
   - Implement pivot tables
   - Add chart support
   - Create data validation rules
   - Support Excel templates

2. **Integration**
   - Add database connectivity
   - Implement web API endpoints
   - Create scheduled tasks

3. **UI Development**
   - Build a web interface
   - Create a desktop application
   - Develop mobile support

---

# ✅ Summary

This project provides a robust foundation for Excel automation in C# using EPPlus. Follow the structure and examples to build your own Excel automation solution.

## 🐳 Docker Support

### Prerequisites
In addition to the existing prerequisites, you'll need:
- [Docker](https://www.docker.com/get-started)
- [Docker Compose](https://docs.docker.com/compose/install/)

### Running with Docker

1. Build the Docker image:
```bash
docker-compose build
```

2. Create a data directory for Excel files:
```bash
mkdir -p data
```

3. Run the application:
```bash
# Create a new Excel file
docker-compose run --rm spreadsheet-editor --file /app/data/myfile.xlsx --new

# Write to a cell
docker-compose run --rm spreadsheet-editor --file /app/data/myfile.xlsx --cell A1 --value "Hello" --write

# Read from a cell
docker-compose run --rm spreadsheet-editor --file /app/data/myfile.xlsx --cell A1 --read
```

### Running Tests
The project includes an automated test script for Docker:

```bash
chmod +x test-docker.sh  # Make the test script executable
./test-docker.sh
```

This test script will:
- Create a new Excel file
- Write a test value to cell A1
- Read the value back and verify it matches
- Clean up the test file

### Docker Project Structure
Additional files for Docker support:
```
SpreadsheetEditor/
├── ...
├── data/                                # Excel file storage (mounted volume)
├── Dockerfile                           # Docker build configuration
├── docker-compose.yml                   # Docker Compose configuration
└── test-docker.sh                       # Docker test script
```
