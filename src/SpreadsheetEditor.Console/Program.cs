using System;
using System.IO;
using System.Threading.Tasks;
using CommandLine;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using OfficeOpenXml;

namespace SpreadsheetEditor.Console
{
    public class Program
    {
        private static IConfiguration Configuration = null!;

        public static async Task Main(string[] args)
        {
            // Set up configuration
            Configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            // Set up dependency injection
            var serviceProvider = new ServiceCollection()
                .AddSingleton(Configuration)
                .BuildServiceProvider();

            // Parse command line arguments
            var result = await Parser.Default.ParseArguments<Options>(args)
                .WithParsedAsync(async options => await RunOptions(options));
            
            if (result.Tag == ParserResultType.NotParsed)
            {
                HandleParseError(((NotParsed<Options>)result).Errors);
            }
        }

        private static async Task RunOptions(Options options)
        {
            try
            {
                // Set EPPlus license context
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                if (options.Create)
                {
                    using (var package = new ExcelPackage())
                    {
                        var sheetName = options.SheetName ?? Configuration["AppSettings:DefaultWorksheetName"] ?? "Sheet1";
                        var worksheet = package.Workbook.Worksheets.Add(sheetName);
                        await package.SaveAsAsync(new FileInfo(options.FilePath));
                        System.Console.WriteLine($"Created new Excel file '{options.FilePath}' with worksheet '{sheetName}'");
                        return;
                    }
                }

                // Validate options for read/write operations
                if (!options.Read && !options.Write)
                {
                    System.Console.WriteLine("Please specify either --read or --write");
                    return;
                }

                if (options.Write && string.IsNullOrEmpty(options.Value))
                {
                    System.Console.WriteLine("Please provide a value to write using --value");
                    return;
                }

                if ((options.Read || options.Write) && string.IsNullOrEmpty(options.CellAddress))
                {
                    System.Console.WriteLine("Please specify a cell address using --cell");
                    return;
                }

                using (var package = new ExcelPackage(new FileInfo(options.FilePath)))
                {
                    var sheetName = options.SheetName ?? Configuration["AppSettings:DefaultWorksheetName"] ?? "Sheet1";
                    var worksheet = package.Workbook.Worksheets[sheetName];
                    if (worksheet == null)
                    {
                        System.Console.WriteLine($"Worksheet '{sheetName}' not found.");
                        return;
                    }

                    if (options.Read)
                    {
                        var value = worksheet.Cells[options.CellAddress!].Value;
                        System.Console.WriteLine($"Value in cell {options.CellAddress}: {value}");
                    }
                    else if (options.Write)
                    {
                        worksheet.Cells[options.CellAddress!].Value = options.Value;
                        await package.SaveAsync();
                        System.Console.WriteLine($"Successfully wrote '{options.Value}' to cell {options.CellAddress}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Console.WriteLine($"Error: {ex.Message}");
            }
        }

        private static void HandleParseError(IEnumerable<Error> errs)
        {
            foreach (var error in errs)
            {
                System.Console.WriteLine($"Error: {error}");
            }
        }
    }
}
