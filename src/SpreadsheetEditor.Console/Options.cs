using CommandLine;

namespace SpreadsheetEditor.Console
{
    public class Options
    {
        [Option('f', "file", Required = true, HelpText = "Path to the Excel file")]
        public required string FilePath { get; set; }

        [Option('s', "sheet", Required = false, HelpText = "Worksheet name (default: Sheet1)")]
        public string? SheetName { get; set; }

        [Option('c', "cell", Required = false, HelpText = "Cell address (e.g., A1)")]
        public string? CellAddress { get; set; }

        [Option('v', "value", Required = false, HelpText = "Value to write to the cell")]
        public string? Value { get; set; }

        [Option('r', "read", Required = false, HelpText = "Read value from cell")]
        public bool Read { get; set; }

        [Option('w', "write", Required = false, HelpText = "Write value to cell")]
        public bool Write { get; set; }

        [Option('n', "new", Required = false, HelpText = "Create a new Excel file")]
        public bool Create { get; set; }
    }
} 