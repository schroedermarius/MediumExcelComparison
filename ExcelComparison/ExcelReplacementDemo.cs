using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;

namespace ExcelComparison;

public class ExcelReplacementDemo
{
    public static void RunDemo()
    {
        // Ensure template exists
        ExcelTemplateGenerator.EnsureTemplateExists();
        Console.WriteLine("=== Excel Replacement Demo ===");
        Console.WriteLine();

        // Collect all variables
        Console.WriteLine("Please enter value for 'Vehicle Registration':");
        var vehicleRegistration = Console.ReadLine();
        Console.WriteLine("Please enter value for 'Dashboard':");
        var dashboard = Console.ReadLine();
        Console.WriteLine("Please enter value for 'Defect Description':");
        var defectDescription = Console.ReadLine();
        Console.WriteLine("Please enter value for 'Revenue Q1':");
        var revenueQ1 = Console.ReadLine();
        Console.WriteLine("Please enter value for 'Profit Q1':");
        var profitQ1 = Console.ReadLine();
        Console.WriteLine("Please enter value for 'Status A':");
        var statusA = Console.ReadLine();
        Console.WriteLine("Please enter value for 'Budget A':");
        var budgetA = Console.ReadLine();

        var fileId = Guid.NewGuid();
        // Demo OpenXML SDK
        DemoOpenXmlSdk(fileId, vehicleRegistration, dashboard, defectDescription, revenueQ1, profitQ1, statusA, budgetA);
        // Demo ClosedXML
        DemoClosedXml(fileId, vehicleRegistration, dashboard, defectDescription, revenueQ1, profitQ1, statusA, budgetA);
        Console.WriteLine();
        Console.WriteLine("Demo completed!");
        Console.WriteLine($"Files generated:");
        Console.WriteLine($"- {fileId}_OpenXML.xlsx");
        Console.WriteLine($"- {fileId}_ClosedXML.xlsx");
    }

    private static void DemoOpenXmlSdk(Guid fileId, string? vehicleRegistration, string? dashboard, string? defectDescription, string? revenueQ1, string? profitQ1, string? statusA, string? budgetA)
    {
        var fileName = $"{fileId}_OpenXML.xlsx";
        File.Copy("Assets/Template.xlsx", fileName, true);
        Console.WriteLine($"Created a new file with ID: {fileName}");
        Console.WriteLine("Using Open XML SDK to replace variables in the Excel file...");
        Stopwatch stopwatch = Stopwatch.StartNew();
        using (FileStream fileStream = new(fileName, FileMode.Open, FileAccess.ReadWrite))
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileStream, true))
        {
            WorkbookPart? workbookPart = document.WorkbookPart;
            WorksheetPart? worksheetPart = workbookPart?.WorksheetParts.First();
            SheetData? sheetData = worksheetPart?.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                Console.WriteLine("No sheet data found in the workbook.");
                return;
            }
            foreach (Row row in sheetData.Elements<Row>())
            {
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellValue is not null && cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                    {
                        int id = int.Parse(cell.CellValue.InnerText);
                        SharedStringItem cellTextItem = workbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                        string text = cellTextItem.InnerText;
                        text = text.Replace("##VehicleRegistration##", vehicleRegistration ?? "DefaultVehicle")
                            .Replace("##Dashboard##", dashboard ?? "DefaultDashboard")
                            .Replace("##DefectDescription##", defectDescription ?? "DefaultDefect")
                            .Replace("##Date##", DateTime.Now.ToString("MM/dd/yyyy"))
                            .Replace("##Revenue_Q1##", revenueQ1 ?? "450,000")
                            .Replace("##Profit_Q1##", profitQ1 ?? "85,000")
                            .Replace("##Costs_Q1##", "365,000")
                            .Replace("##Margin_Q1##", "18.9")
                            .Replace("##Revenue_Q2##", "520,000")
                            .Replace("##Profit_Q2##", "95,000")
                            .Replace("##Costs_Q2##", "425,000")
                            .Replace("##Margin_Q2##", "18.3")
                            .Replace("##Revenue_Q3##", "580,000")
                            .Replace("##Profit_Q3##", "110,000")
                            .Replace("##Costs_Q3##", "470,000")
                            .Replace("##Margin_Q3##", "19.0")
                            .Replace("##Revenue_Q4##", "620,000")
                            .Replace("##Profit_Q4##", "125,000")
                            .Replace("##Costs_Q4##", "495,000")
                            .Replace("##Margin_Q4##", "20.2")
                            .Replace("##Status_A##", statusA ?? "Completed")
                            .Replace("##Budget_A##", budgetA ?? "75,000")
                            .Replace("##Status_B##", "In Progress")
                            .Replace("##Budget_B##", "120,000")
                            .Replace("##Status_C##", "Planned")
                            .Replace("##Budget_C##", "200,000");
                        cellTextItem.GetFirstChild<Text>()!.Text = text;
                    }
                }
            }
            worksheetPart?.Worksheet.Save();
        }
        stopwatch.Stop();
        Console.WriteLine($"Replacement completed in {stopwatch.ElapsedMilliseconds} ms using Open XML SDK!");
    }

    private static void DemoClosedXml(Guid fileId, string? vehicleRegistration, string? dashboard, string? defectDescription, string? revenueQ1, string? profitQ1, string? statusA, string? budgetA)
    {
        var fileName = $"{fileId}_ClosedXML.xlsx";
        File.Copy("Assets/Template.xlsx", fileName, true);
        Console.WriteLine($"Created a new file with ID: {fileName}");
        Console.WriteLine("Using ClosedXML to replace variables in the Excel file...");
        Stopwatch stopwatch = Stopwatch.StartNew();
        using (var workbook = new XLWorkbook(fileName))
        {
            var worksheet = workbook.Worksheet(1);
            foreach (var cell in worksheet.CellsUsed())
            {
                if (cell.HasFormula) continue;
                var value = cell.GetString();
                value = value.Replace("##VehicleRegistration##", vehicleRegistration ?? "DefaultVehicle")
                             .Replace("##Dashboard##", dashboard ?? "DefaultDashboard")
                             .Replace("##DefectDescription##", defectDescription ?? "DefaultDefect")
                             .Replace("##Date##", DateTime.Now.ToString("MM/dd/yyyy"))
                             .Replace("##Revenue_Q1##", revenueQ1 ?? "450,000")
                             .Replace("##Profit_Q1##", profitQ1 ?? "85,000")
                             .Replace("##Costs_Q1##", "365,000")
                             .Replace("##Margin_Q1##", "18.9")
                             .Replace("##Revenue_Q2##", "520,000")
                             .Replace("##Profit_Q2##", "95,000")
                             .Replace("##Costs_Q2##", "425,000")
                             .Replace("##Margin_Q2##", "18.3")
                             .Replace("##Revenue_Q3##", "580,000")
                             .Replace("##Profit_Q3##", "110,000")
                             .Replace("##Costs_Q3##", "470,000")
                             .Replace("##Margin_Q3##", "19.0")
                             .Replace("##Revenue_Q4##", "620,000")
                             .Replace("##Profit_Q4##", "125,000")
                             .Replace("##Costs_Q4##", "495,000")
                             .Replace("##Margin_Q4##", "20.2")
                             .Replace("##Status_A##", statusA ?? "Completed")
                             .Replace("##Budget_A##", budgetA ?? "75,000")
                             .Replace("##Status_B##", "In Progress")
                             .Replace("##Budget_B##", "120,000")
                             .Replace("##Status_C##", "Planned")
                             .Replace("##Budget_C##", "200,000");
                cell.Value = value;
            }
            workbook.Save();
        }
        stopwatch.Stop();
        Console.WriteLine($"Replacement completed in {stopwatch.ElapsedMilliseconds} ms using ClosedXML!");
    }
}
