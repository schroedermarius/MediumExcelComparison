using BenchmarkDotNet.Running;
using ExcelComparison;

Console.WriteLine("Excel Replacement Comparison Tool");
Console.WriteLine("=================================");
Console.WriteLine();
Console.WriteLine("Choose an option:");
Console.WriteLine("1. Run Demo (Interactive Excel replacement)");
Console.WriteLine("2. Run Benchmarks (Performance comparison)");
Console.WriteLine();
Console.Write("Enter your choice (1 or 2): ");

var choice = Console.ReadLine();

switch (choice)
{
    case "1":
        ExcelReplacementDemo.RunDemo();
        break;
    case "2":
        Console.WriteLine("Starting benchmarks...");
        Console.WriteLine("This may take a few minutes to complete.");
        Console.WriteLine();
        BenchmarkRunner.Run<ExcelBenchmark>();
        break;
    default:
        Console.WriteLine("Invalid choice. Running demo by default...");
        ExcelReplacementDemo.RunDemo();
        break;
}

Console.WriteLine("Press any key to exit...");
Console.ReadKey();
