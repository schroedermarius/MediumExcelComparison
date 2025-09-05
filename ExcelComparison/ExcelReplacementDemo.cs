using System.Diagnostics;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;

namespace ExcelComparison;

public class ExcelReplacementDemo
{
    public static void RunDemo()
    {
        // Sicherstellen, dass die Vorlage existiert
        ExcelTemplateGenerator.EnsureTemplateExists();
        Console.WriteLine("=== Excel Replacement Demo ===");
        Console.WriteLine();

        // Alle Variablen abfragen
        Console.WriteLine("Bitte Wert für 'Fahrzeugschein' eingeben:");
        var fahrzeugschein = Console.ReadLine();
        Console.WriteLine("Bitte Wert für 'Armaturen' eingeben:");
        var armaturen = Console.ReadLine();
        Console.WriteLine("Bitte Wert für 'MangelbeschreibungSB' eingeben:");
        var mangelbeschreibungSb = Console.ReadLine();
        Console.WriteLine("Bitte Wert für 'Umsatz Q1' eingeben:");
        var umsatzQ1 = Console.ReadLine();
        Console.WriteLine("Bitte Wert für 'Gewinn Q1' eingeben:");
        var gewinnQ1 = Console.ReadLine();
        Console.WriteLine("Bitte Wert für 'Status A' eingeben:");
        var statusA = Console.ReadLine();
        Console.WriteLine("Bitte Wert für 'Budget A' eingeben:");
        var budgetA = Console.ReadLine();

        var fileId = Guid.NewGuid();
        // Demo OpenXML SDK
        DemoOpenXmlSdk(fileId, fahrzeugschein, armaturen, mangelbeschreibungSb, umsatzQ1, gewinnQ1, statusA, budgetA);
        // Demo ClosedXML
        DemoClosedXml(fileId, fahrzeugschein, armaturen, mangelbeschreibungSb, umsatzQ1, gewinnQ1, statusA, budgetA);
        Console.WriteLine();
        Console.WriteLine("Demo abgeschlossen!");
        Console.WriteLine($"Dateien generiert:");
        Console.WriteLine($"- {fileId}_OpenXML.xlsx");
        Console.WriteLine($"- {fileId}_ClosedXML.xlsx");
    }

    private static void DemoOpenXmlSdk(Guid fileId, string? fahrzeugschein, string? armaturen, string? mangelbeschreibungSb, string? umsatzQ1, string? gewinnQ1, string? statusA, string? budgetA)
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
                        text = text.Replace("##Fahrzeugschein##", fahrzeugschein ?? "Default1")
                            .Replace("##Armaturen##", armaturen ?? "Default2")
                            .Replace("##MangelbeschreibungSB##", mangelbeschreibungSb ?? "Default3")
                            .Replace("##Datum##", DateTime.Now.ToString("dd.MM.yyyy"))
                            .Replace("##Umsatz_Q1##", umsatzQ1 ?? "450.000")
                            .Replace("##Gewinn_Q1##", gewinnQ1 ?? "85.000")
                            .Replace("##Kosten_Q1##", "365.000")
                            .Replace("##Marge_Q1##", "18.9")
                            .Replace("##Umsatz_Q2##", "520.000")
                            .Replace("##Gewinn_Q2##", "95.000")
                            .Replace("##Kosten_Q2##", "425.000")
                            .Replace("##Marge_Q2##", "18.3")
                            .Replace("##Umsatz_Q3##", "580.000")
                            .Replace("##Gewinn_Q3##", "110.000")
                            .Replace("##Kosten_Q3##", "470.000")
                            .Replace("##Marge_Q3##", "19.0")
                            .Replace("##Umsatz_Q4##", "620.000")
                            .Replace("##Gewinn_Q4##", "125.000")
                            .Replace("##Kosten_Q4##", "495.000")
                            .Replace("##Marge_Q4##", "20.2")
                            .Replace("##Status_A##", statusA ?? "Abgeschlossen")
                            .Replace("##Budget_A##", budgetA ?? "75.000")
                            .Replace("##Status_B##", "In Bearbeitung")
                            .Replace("##Budget_B##", "120.000")
                            .Replace("##Status_C##", "Geplant")
                            .Replace("##Budget_C##", "200.000");
                        cellTextItem.GetFirstChild<Text>()!.Text = text;
                    }
                }
            }
            worksheetPart?.Worksheet.Save();
        }
        stopwatch.Stop();
        Console.WriteLine($"Replacement completed in {stopwatch.ElapsedMilliseconds} ms using Open XML SDK!");
    }

    private static void DemoClosedXml(Guid fileId, string? fahrzeugschein, string? armaturen, string? mangelbeschreibungSb, string? umsatzQ1, string? gewinnQ1, string? statusA, string? budgetA)
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
                value = value.Replace("##Fahrzeugschein##", fahrzeugschein ?? "Default1")
                             .Replace("##Armaturen##", armaturen ?? "Default2")
                             .Replace("##MangelbeschreibungSB##", mangelbeschreibungSb ?? "Default3")
                             .Replace("##Datum##", DateTime.Now.ToString("dd.MM.yyyy"))
                             .Replace("##Umsatz_Q1##", umsatzQ1 ?? "450.000")
                             .Replace("##Gewinn_Q1##", gewinnQ1 ?? "85.000")
                             .Replace("##Kosten_Q1##", "365.000")
                             .Replace("##Marge_Q1##", "18.9")
                             .Replace("##Umsatz_Q2##", "520.000")
                             .Replace("##Gewinn_Q2##", "95.000")
                             .Replace("##Kosten_Q2##", "425.000")
                             .Replace("##Marge_Q2##", "18.3")
                             .Replace("##Umsatz_Q3##", "580.000")
                             .Replace("##Gewinn_Q3##", "110.000")
                             .Replace("##Kosten_Q3##", "470.000")
                             .Replace("##Marge_Q3##", "19.0")
                             .Replace("##Umsatz_Q4##", "620.000")
                             .Replace("##Gewinn_Q4##", "125.000")
                             .Replace("##Kosten_Q4##", "495.000")
                             .Replace("##Marge_Q4##", "20.2")
                             .Replace("##Status_A##", statusA ?? "Abgeschlossen")
                             .Replace("##Budget_A##", budgetA ?? "75.000")
                             .Replace("##Status_B##", "In Bearbeitung")
                             .Replace("##Budget_B##", "120.000")
                             .Replace("##Status_C##", "Geplant")
                             .Replace("##Budget_C##", "200.000");
                cell.Value = value;
            }
            workbook.Save();
        }
        stopwatch.Stop();
        Console.WriteLine($"Replacement completed in {stopwatch.ElapsedMilliseconds} ms using ClosedXML!");
    }
}
