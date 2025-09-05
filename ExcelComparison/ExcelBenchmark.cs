using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using BenchmarkDotNet.Jobs;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;

namespace ExcelComparison;

[Config(typeof(Config))]
[MemoryDiagnoser]
[SimpleJob(RuntimeMoniker.Net90)]
public class ExcelBenchmark
{
    private class Config : ManualConfig
    {
        public Config()
        {
            AddJob(Job.Default.WithWarmupCount(3).WithIterationCount(5));
        }
    }

    private string _firmenname = "TechSolutions GmbH";
    private string _geschäftsführer = "Max Mustermann";
    private string _standort = "München, Deutschland";
    private string _mitarbeiter = "150";
    private string _jahr = "2025";
    private string _umsatzQ1 = "450.000";
    private string _gewinnQ1 = "85.000";
    private string _kostenQ1 = "365.000";
    private string _margeQ1 = "18.9";
    private string _umsatzQ2 = "520.000";
    private string _gewinnQ2 = "95.000";
    private string _kostenQ2 = "425.000";
    private string _margeQ2 = "18.3";
    private string _umsatzQ3 = "580.000";
    private string _gewinnQ3 = "110.000";
    private string _kostenQ3 = "470.000";
    private string _margeQ3 = "19.0";
    private string _umsatzQ4 = "620.000";
    private string _gewinnQ4 = "125.000";
    private string _kostenQ4 = "495.000";
    private string _margeQ4 = "20.2";
    private string _statusA = "Abgeschlossen";
    private string _budgetA = "75.000";
    private string _statusB = "In Bearbeitung";
    private string _budgetB = "120.000";
    private string _statusC = "Geplant";
    private string _budgetC = "200.000";
    private string _bemerkungen = "Sehr erfolgreiches Quartal mit überdurchschnittlichem Wachstum. Neue Produktlinie wurde erfolgreich eingeführt.";
    private string _templatePath = "Assets/Template.xlsx";
    
    [GlobalSetup]
    public void Setup()
    {
        // Sicherstellen, dass die Vorlage existiert
        ExcelTemplateGenerator.EnsureTemplateExists();
        
        // Ensure template file exists
        if (!File.Exists(_templatePath))
        {
            throw new FileNotFoundException($"Template file not found: {_templatePath}");
        }
    }

    [Benchmark(Baseline = true)]
    public void OpenXmlSdk()
    {
        var fileName = $"{Guid.NewGuid()}_OpenXML_Benchmark.xlsx";
        try
        {
            File.Copy(_templatePath, fileName, true);

            using FileStream fileStream = new(fileName, FileMode.Open, FileAccess.ReadWrite);
            using SpreadsheetDocument document = SpreadsheetDocument.Open(fileStream, true);

            WorkbookPart? workbookPart = document.WorkbookPart;
            WorksheetPart? worksheetPart = workbookPart?.WorksheetParts.First();
            SheetData? sheetData = worksheetPart?.Worksheet.GetFirstChild<SheetData>();

            if (sheetData == null) return;

            foreach (Row row in sheetData.Elements<Row>())
            {
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellValue is not null && cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                    {
                        int id = int.Parse(cell.CellValue.InnerText);
                        SharedStringItem cellTextItem = workbookPart!.SharedStringTablePart!.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                        string text = cellTextItem.InnerText;

                        text = text.Replace("##Firmenname##", _firmenname)
                            .Replace("##Geschäftsführer##", _geschäftsführer)
                            .Replace("##Standort##", _standort)
                            .Replace("##Mitarbeiter##", _mitarbeiter)
                            .Replace("##Jahr##", _jahr)
                            .Replace("##Umsatz_Q1##", _umsatzQ1)
                            .Replace("##Gewinn_Q1##", _gewinnQ1)
                            .Replace("##Kosten_Q1##", _kostenQ1)
                            .Replace("##Marge_Q1##", _margeQ1)
                            .Replace("##Umsatz_Q2##", _umsatzQ2)
                            .Replace("##Gewinn_Q2##", _gewinnQ2)
                            .Replace("##Kosten_Q2##", _kostenQ2)
                            .Replace("##Marge_Q2##", _margeQ2)
                            .Replace("##Umsatz_Q3##", _umsatzQ3)
                            .Replace("##Gewinn_Q3##", _gewinnQ3)
                            .Replace("##Kosten_Q3##", _kostenQ3)
                            .Replace("##Marge_Q3##", _margeQ3)
                            .Replace("##Umsatz_Q4##", _umsatzQ4)
                            .Replace("##Gewinn_Q4##", _gewinnQ4)
                            .Replace("##Kosten_Q4##", _kostenQ4)
                            .Replace("##Marge_Q4##", _margeQ4)
                            .Replace("##Status_A##", _statusA)
                            .Replace("##Budget_A##", _budgetA)
                            .Replace("##Status_B##", _statusB)
                            .Replace("##Budget_B##", _budgetB)
                            .Replace("##Status_C##", _statusC)
                            .Replace("##Budget_C##", _budgetC)
                            .Replace("##Bemerkungen##", _bemerkungen)
                            .Replace("##Datum##", DateTime.Now.ToString("dd.MM.yyyy"));

                        cellTextItem.GetFirstChild<Text>()!.Text = text;
                    }
                }
            }

            worksheetPart?.Worksheet.Save();
        }
        finally
        {
            if (File.Exists(fileName))
                File.Delete(fileName);
        }
    }

    [Benchmark]
    public void ClosedXml()
    {
        var fileName = $"{Guid.NewGuid()}_ClosedXML_Benchmark.xlsx";
        try
        {
            File.Copy(_templatePath, fileName, true);

            using var workbook = new XLWorkbook(fileName);
            var worksheet = workbook.Worksheet(1);

            foreach (var cell in worksheet.CellsUsed())
            {
                if (cell.HasFormula)
                {
                    continue; // Skip cells with formulas
                }

                var value = cell.GetString();
                value = value.Replace("##Firmenname##", _firmenname)
                             .Replace("##Geschäftsführer##", _geschäftsführer)
                             .Replace("##Standort##", _standort)
                             .Replace("##Mitarbeiter##", _mitarbeiter)
                             .Replace("##Jahr##", _jahr)
                             .Replace("##Umsatz_Q1##", _umsatzQ1)
                             .Replace("##Gewinn_Q1##", _gewinnQ1)
                             .Replace("##Kosten_Q1##", _kostenQ1)
                             .Replace("##Marge_Q1##", _margeQ1)
                             .Replace("##Umsatz_Q2##", _umsatzQ2)
                             .Replace("##Gewinn_Q2##", _gewinnQ2)
                             .Replace("##Kosten_Q2##", _kostenQ2)
                             .Replace("##Marge_Q2##", _margeQ2)
                             .Replace("##Umsatz_Q3##", _umsatzQ3)
                             .Replace("##Gewinn_Q3##", _gewinnQ3)
                             .Replace("##Kosten_Q3##", _kostenQ3)
                             .Replace("##Marge_Q3##", _margeQ3)
                             .Replace("##Umsatz_Q4##", _umsatzQ4)
                             .Replace("##Gewinn_Q4##", _gewinnQ4)
                             .Replace("##Kosten_Q4##", _kostenQ4)
                             .Replace("##Marge_Q4##", _margeQ4)
                             .Replace("##Status_A##", _statusA)
                             .Replace("##Budget_A##", _budgetA)
                             .Replace("##Status_B##", _statusB)
                             .Replace("##Budget_B##", _budgetB)
                             .Replace("##Status_C##", _statusC)
                             .Replace("##Budget_C##", _budgetC)
                             .Replace("##Bemerkungen##", _bemerkungen)
                             .Replace("##Datum##", DateTime.Now.ToString("dd.MM.yyyy"));

                cell.Value = value;
            }

            workbook.Save();
        }
        finally
        {
            if (File.Exists(fileName))
                File.Delete(fileName);
        }
    }
}
