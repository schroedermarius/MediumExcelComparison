using ClosedXML.Excel;

namespace ExcelComparison
{
    public static class ExcelTemplateGenerator
    {
        public static void EnsureTemplateExists()
        {
            var path = "Assets/Template.xlsx";
            if (File.Exists(path)) return;

            Directory.CreateDirectory("Assets");
            using (var wb = new XLWorkbook())
            {
                var ws = wb.AddWorksheet("Data");
                
                // Demo placeholders
                ws.Cell(1, 1).Value = "Vehicle Registration:";
                ws.Cell(1, 2).Value = "##VehicleRegistration##";
                ws.Cell(2, 1).Value = "Dashboard:";
                ws.Cell(2, 2).Value = "##Dashboard##";
                ws.Cell(3, 1).Value = "Defect Description:";
                ws.Cell(3, 2).Value = "##DefectDescription##";
                ws.Cell(4, 1).Value = "Date:";
                ws.Cell(4, 2).Value = "##Date##";
                
                // Benchmark placeholders (Company information)
                ws.Cell(1, 4).Value = "Company Name:";
                ws.Cell(1, 5).Value = "##CompanyName##";
                ws.Cell(2, 4).Value = "CEO:";
                ws.Cell(2, 5).Value = "##CEO##";
                ws.Cell(3, 4).Value = "Location:";
                ws.Cell(3, 5).Value = "##Location##";
                ws.Cell(4, 4).Value = "Employees:";
                ws.Cell(4, 5).Value = "##Employees##";
                ws.Cell(5, 4).Value = "Year:";
                ws.Cell(5, 5).Value = "##Year##";
                ws.Cell(6, 4).Value = "Remarks:";
                ws.Cell(6, 5).Value = "##Remarks##";
                
                // Quarterly overview
                ws.Cell(6, 1).Value = "Quarter";
                ws.Cell(6, 2).Value = "Revenue";
                ws.Cell(6, 3).Value = "Profit";
                ws.Cell(6, 4).Value = "Costs";
                ws.Cell(6, 5).Value = "Margin (%)";
                ws.Cell(7, 1).Value = "Q1";
                ws.Cell(7, 2).Value = "##Revenue_Q1##";
                ws.Cell(7, 3).Value = "##Profit_Q1##";
                ws.Cell(7, 4).Value = "##Costs_Q1##";
                ws.Cell(7, 5).Value = "##Margin_Q1##";
                ws.Cell(8, 1).Value = "Q2";
                ws.Cell(8, 2).Value = "##Revenue_Q2##";
                ws.Cell(8, 3).Value = "##Profit_Q2##";
                ws.Cell(8, 4).Value = "##Costs_Q2##";
                ws.Cell(8, 5).Value = "##Margin_Q2##";
                ws.Cell(9, 1).Value = "Q3";
                ws.Cell(9, 2).Value = "##Revenue_Q3##";
                ws.Cell(9, 3).Value = "##Profit_Q3##";
                ws.Cell(9, 4).Value = "##Costs_Q3##";
                ws.Cell(9, 5).Value = "##Margin_Q3##";
                ws.Cell(10, 1).Value = "Q4";
                ws.Cell(10, 2).Value = "##Revenue_Q4##";
                ws.Cell(10, 3).Value = "##Profit_Q4##";
                ws.Cell(10, 4).Value = "##Costs_Q4##";
                ws.Cell(10, 5).Value = "##Margin_Q4##";
                
                // Project status
                ws.Cell(12, 1).Value = "Project";
                ws.Cell(12, 2).Value = "Status";
                ws.Cell(12, 3).Value = "Budget";
                ws.Cell(13, 1).Value = "A";
                ws.Cell(13, 2).Value = "##Status_A##";
                ws.Cell(13, 3).Value = "##Budget_A##";
                ws.Cell(14, 1).Value = "B";
                ws.Cell(14, 2).Value = "##Status_B##";
                ws.Cell(14, 3).Value = "##Budget_B##";
                ws.Cell(15, 1).Value = "C";
                ws.Cell(15, 2).Value = "##Status_C##";
                ws.Cell(15, 3).Value = "##Budget_C##";
                
                wb.SaveAs(path);
            }
        }
    }
}
