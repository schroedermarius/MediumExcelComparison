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

    private string _companyName = "TechSolutions Inc";
    private string _ceo = "John Smith";
    private string _location = "Munich, Germany";
    private string _employees = "150";
    private string _year = "2025";
    private string _revenueQ1 = "450,000";
    private string _profitQ1 = "85,000";
    private string _costsQ1 = "365,000";
    private string _marginQ1 = "18.9";
    private string _revenueQ2 = "520,000";
    private string _profitQ2 = "95,000";
    private string _costsQ2 = "425,000";
    private string _marginQ2 = "18.3";
    private string _revenueQ3 = "580,000";
    private string _profitQ3 = "110,000";
    private string _costsQ3 = "470,000";
    private string _marginQ3 = "19.0";
    private string _revenueQ4 = "620,000";
    private string _profitQ4 = "125,000";
    private string _costsQ4 = "495,000";
    private string _marginQ4 = "20.2";
    private string _statusA = "Completed";
    private string _budgetA = "75,000";
    private string _statusB = "In Progress";
    private string _budgetB = "120,000";
    private string _statusC = "Planned";
    private string _budgetC = "200,000";
    private string _remarks = "Very successful quarter with above-average growth. New product line was successfully launched.";
    private string _templatePath = "Assets/Template.xlsx";
    
    [GlobalSetup]
    public void Setup()
    {
        // Ensure template exists
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

                        text = text.Replace("##CompanyName##", _companyName)
                            .Replace("##CEO##", _ceo)
                            .Replace("##Location##", _location)
                            .Replace("##Employees##", _employees)
                            .Replace("##Year##", _year)
                            .Replace("##Revenue_Q1##", _revenueQ1)
                            .Replace("##Profit_Q1##", _profitQ1)
                            .Replace("##Costs_Q1##", _costsQ1)
                            .Replace("##Margin_Q1##", _marginQ1)
                            .Replace("##Revenue_Q2##", _revenueQ2)
                            .Replace("##Profit_Q2##", _profitQ2)
                            .Replace("##Costs_Q2##", _costsQ2)
                            .Replace("##Margin_Q2##", _marginQ2)
                            .Replace("##Revenue_Q3##", _revenueQ3)
                            .Replace("##Profit_Q3##", _profitQ3)
                            .Replace("##Costs_Q3##", _costsQ3)
                            .Replace("##Margin_Q3##", _marginQ3)
                            .Replace("##Revenue_Q4##", _revenueQ4)
                            .Replace("##Profit_Q4##", _profitQ4)
                            .Replace("##Costs_Q4##", _costsQ4)
                            .Replace("##Margin_Q4##", _marginQ4)
                            .Replace("##Status_A##", _statusA)
                            .Replace("##Budget_A##", _budgetA)
                            .Replace("##Status_B##", _statusB)
                            .Replace("##Budget_B##", _budgetB)
                            .Replace("##Status_C##", _statusC)
                            .Replace("##Budget_C##", _budgetC)
                            .Replace("##Remarks##", _remarks)
                            .Replace("##Date##", DateTime.Now.ToString("MM/dd/yyyy"));

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
                value = value.Replace("##CompanyName##", _companyName)
                             .Replace("##CEO##", _ceo)
                             .Replace("##Location##", _location)
                             .Replace("##Employees##", _employees)
                             .Replace("##Year##", _year)
                             .Replace("##Revenue_Q1##", _revenueQ1)
                             .Replace("##Profit_Q1##", _profitQ1)
                             .Replace("##Costs_Q1##", _costsQ1)
                             .Replace("##Margin_Q1##", _marginQ1)
                             .Replace("##Revenue_Q2##", _revenueQ2)
                             .Replace("##Profit_Q2##", _profitQ2)
                             .Replace("##Costs_Q2##", _costsQ2)
                             .Replace("##Margin_Q2##", _marginQ2)
                             .Replace("##Revenue_Q3##", _revenueQ3)
                             .Replace("##Profit_Q3##", _profitQ3)
                             .Replace("##Costs_Q3##", _costsQ3)
                             .Replace("##Margin_Q3##", _marginQ3)
                             .Replace("##Revenue_Q4##", _revenueQ4)
                             .Replace("##Profit_Q4##", _profitQ4)
                             .Replace("##Costs_Q4##", _costsQ4)
                             .Replace("##Margin_Q4##", _marginQ4)
                             .Replace("##Status_A##", _statusA)
                             .Replace("##Budget_A##", _budgetA)
                             .Replace("##Status_B##", _statusB)
                             .Replace("##Budget_B##", _budgetB)
                             .Replace("##Status_C##", _statusC)
                             .Replace("##Budget_C##", _budgetC)
                             .Replace("##Remarks##", _remarks)
                             .Replace("##Date##", DateTime.Now.ToString("MM/dd/yyyy"));

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
