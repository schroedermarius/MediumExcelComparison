# 📊 Excel Template Manipulation in .NET: OpenXML SDK vs ClosedXML

> **Fast or friendly?** A comprehensive comparison between raw speed (Open XML SDK) and developer joy (ClosedXML) for Excel template manipulation in .NET.

## 🎯 Overview

This project demonstrates and benchmarks two popular .NET libraries for Excel manipulation when working with template-based document generation. We evaluate both approaches for **ergonomics** (developer experience) and **performance** (using BenchmarkDotNet).

### The Contenders

- **[Open XML SDK](https://github.com/dotnet/Open-XML-SDK)** (`DocumentFormat.OpenXml` v3.3.0) - Microsoft's low-level, high-performance library
- **[ClosedXML](https://github.com/ClosedXML/ClosedXML)** (`ClosedXML` v0.105.0) - Community-driven, developer-friendly wrapper

## 🚀 Quick Start

### Prerequisites

- .NET 9.0 or later
- An Excel template file (automatically generated if missing)

### Running the Demo

```bash
git clone <repository-url>
cd ExcelComparison
dotnet run
```

Choose from the interactive menu:
1. **Demo Mode** - Interactive Excel replacement with user input
2. **Benchmark Mode** - Performance comparison using BenchmarkDotNet

## 🧩 What This Demo Does

The application demonstrates a common real-world scenario: taking form data and injecting it into an existing Excel template by replacing placeholder tokens.

### Template Processing Flow

1. **Template Validation** - Ensures `Assets/Template.xlsx` exists (generates if missing)
2. **User Input Collection** - Prompts for values like:
   - `Fahrzeugschein` (Vehicle Registration)
   - `Umsatz_Q1` (Q1 Revenue)
   - `Status_A` (Status A)
   - `Budget_A` (Budget A)
3. **Dual Processing** - Creates two copies of the template and processes each with:
   - **Open XML SDK** - Low-level XLSX package manipulation
   - **ClosedXML** - High-level worksheet/cell API
4. **Token Replacement** - Replaces placeholders like `##Umsatz_Q1##` with actual values

## 🔍 Implementation Approaches

### Open XML SDK - Low-Level, Performant

```csharp
// Opens XLSX package, manipulates SharedStringTable directly
using var document = SpreadsheetDocument.Open(filePath, true);
var sharedStringTable = document.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
// Precise control, minimal allocations, more verbose code
```

**Characteristics:**
- ✅ Highest performance and minimal memory allocation
- ✅ Fine-grained control over Excel document structure
- ❌ More verbose, requires deep understanding of OOXML format
- ❌ More development time for complex scenarios

### ClosedXML - High-Level, Ergonomic

```csharp
// Loads workbook, iterates cells with simple API
using var workbook = new XLWorkbook(templatePath);
var worksheet = workbook.Worksheets.First();
foreach (var cell in worksheet.CellsUsed())
{
    // Simple, readable cell manipulation
}
```

**Characteristics:**
- ✅ Developer-friendly API, rapid development
- ✅ Excellent for complex formatting and formulas
- ✅ Great documentation and community support
- ❌ Higher memory usage and slower performance
- ❌ Additional abstraction layer overhead

## 📊 Performance Benchmarks

Benchmarks were conducted using BenchmarkDotNet with the following key results:

| Method     | Runtime | Mean     | Gen0      | Allocated | Alloc Ratio |
|------------|---------|----------|-----------|-----------|-------------|
| OpenXmlSdk | .NET 9.0| 1.057 ms | 50.7813   | 316.62 KB | 1.00        |
| ClosedXml  | .NET 9.0| 3.264 ms | 171.8750  | 1099.18 KB| 3.47        |

### Key Findings

- **OpenXML SDK** is consistently **~3x faster** in execution time
- **OpenXML SDK** uses **~3.4x less memory** allocation
- Performance gap becomes more significant with larger files or higher throughput requirements

## 🎯 Recommendations

### Choose **Open XML SDK** when:

- 🚀 **Performance is critical** - High throughput, many concurrent files
- 💾 **Memory constraints** - Limited memory environment or large files
- 🏢 **Enterprise scale** - Processing thousands of documents
- 🔧 **Team has expertise** - Developers comfortable with low-level APIs

### Choose **ClosedXML** when:

- ⚡ **Development speed matters** - Rapid prototyping, quick implementations
- 🎨 **Complex formatting** - Rich Excel features, charts, conditional formatting
- 👥 **Team productivity** - Mixed skill levels, maintainability priority
- 📊 **Moderate scale** - Hundreds of files, not performance-critical

### Hybrid Approach

Consider starting with **ClosedXML** for rapid development, then optimize critical paths with **OpenXML SDK** if performance becomes a bottleneck.

## 🛠️ Project Structure

```
ExcelComparison/
├── Assets/
│   └── Template.xlsx              # Excel template with placeholders
├── ExcelBenchmark.cs             # BenchmarkDotNet performance tests
├── ExcelReplacementDemo.cs       # Interactive demonstration
├── ExcelTemplateGenerator.cs     # Template file generation
├── Program.cs                    # Main application entry point
└── ExcelComparison.csproj        # Project dependencies
```

## 🔧 Key Classes

- **`ExcelReplacementDemo`** - Interactive demonstration with user input
- **`ExcelBenchmark`** - BenchmarkDotNet performance testing
- **`ExcelTemplateGenerator`** - Creates sample Excel templates with placeholders

## 📦 Dependencies

```xml
<PackageReference Include="ClosedXML" Version="0.105.0" />
<PackageReference Include="DocumentFormat.OpenXml" Version="3.3.0" />
<PackageReference Include="BenchmarkDotNet" Version="0.14.0" />
```

## 🧪 Running Benchmarks

To run comprehensive performance benchmarks:

```bash
dotnet run -c Release
# Select option 2 for benchmarks
```

This will generate detailed benchmark reports including:
- Execution time statistics
- Memory allocation analysis
- Garbage collection metrics
- Statistical significance tests

## 📈 Use Cases

This comparison is particularly valuable for:

- **Document Generation Services** - APIs that generate reports from templates
- **Data Export Systems** - Converting database records to Excel reports
- **Automated Reporting** - Scheduled report generation with dynamic data
- **Template Processing Pipelines** - Batch processing of Excel templates

## 🤝 Contributing

Contributions are welcome! Areas for improvement:
- Additional Excel libraries comparison (EPPlus, NPOI)
- More complex template scenarios
- Different file size benchmarks
- Cross-platform performance analysis

## 📄 License

This project is provided as-is for educational and comparison purposes.

---

*This comparison was conducted to help developers make informed decisions when choosing Excel manipulation libraries for .NET projects. Results may vary based on specific use cases, file sizes, and system configurations.*
