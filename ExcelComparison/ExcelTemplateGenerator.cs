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
                var ws = wb.AddWorksheet("Daten");
                
                // Demo-Platzhalter
                ws.Cell(1, 1).Value = "Fahrzeugschein:";
                ws.Cell(1, 2).Value = "##Fahrzeugschein##";
                ws.Cell(2, 1).Value = "Armaturen:";
                ws.Cell(2, 2).Value = "##Armaturen##";
                ws.Cell(3, 1).Value = "MangelbeschreibungSB:";
                ws.Cell(3, 2).Value = "##MangelbeschreibungSB##";
                ws.Cell(4, 1).Value = "Datum:";
                ws.Cell(4, 2).Value = "##Datum##";
                
                // Benchmark-Platzhalter (Firmeninformationen)
                ws.Cell(1, 4).Value = "Firmenname:";
                ws.Cell(1, 5).Value = "##Firmenname##";
                ws.Cell(2, 4).Value = "Geschäftsführer:";
                ws.Cell(2, 5).Value = "##Geschäftsführer##";
                ws.Cell(3, 4).Value = "Standort:";
                ws.Cell(3, 5).Value = "##Standort##";
                ws.Cell(4, 4).Value = "Mitarbeiter:";
                ws.Cell(4, 5).Value = "##Mitarbeiter##";
                ws.Cell(5, 4).Value = "Jahr:";
                ws.Cell(5, 5).Value = "##Jahr##";
                ws.Cell(6, 4).Value = "Bemerkungen:";
                ws.Cell(6, 5).Value = "##Bemerkungen##";
                
                // Quartalsübersicht
                ws.Cell(6, 1).Value = "Quartal";
                ws.Cell(6, 2).Value = "Umsatz";
                ws.Cell(6, 3).Value = "Gewinn";
                ws.Cell(6, 4).Value = "Kosten";
                ws.Cell(6, 5).Value = "Marge (%)";
                ws.Cell(7, 1).Value = "Q1";
                ws.Cell(7, 2).Value = "##Umsatz_Q1##";
                ws.Cell(7, 3).Value = "##Gewinn_Q1##";
                ws.Cell(7, 4).Value = "##Kosten_Q1##";
                ws.Cell(7, 5).Value = "##Marge_Q1##";
                ws.Cell(8, 1).Value = "Q2";
                ws.Cell(8, 2).Value = "##Umsatz_Q2##";
                ws.Cell(8, 3).Value = "##Gewinn_Q2##";
                ws.Cell(8, 4).Value = "##Kosten_Q2##";
                ws.Cell(8, 5).Value = "##Marge_Q2##";
                ws.Cell(9, 1).Value = "Q3";
                ws.Cell(9, 2).Value = "##Umsatz_Q3##";
                ws.Cell(9, 3).Value = "##Gewinn_Q3##";
                ws.Cell(9, 4).Value = "##Kosten_Q3##";
                ws.Cell(9, 5).Value = "##Marge_Q3##";
                ws.Cell(10, 1).Value = "Q4";
                ws.Cell(10, 2).Value = "##Umsatz_Q4##";
                ws.Cell(10, 3).Value = "##Gewinn_Q4##";
                ws.Cell(10, 4).Value = "##Kosten_Q4##";
                ws.Cell(10, 5).Value = "##Marge_Q4##";
                
                // Projektstatus
                ws.Cell(12, 1).Value = "Projekt";
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
