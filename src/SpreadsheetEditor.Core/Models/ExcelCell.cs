namespace SpreadsheetEditor.Core.Models
{
    public class ExcelCell
    {
        public string Address { get; set; } = string.Empty;
        public object? Value { get; set; }
        public string? Formula { get; set; }
    }
} 