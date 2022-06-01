using ClosedXML.Excel;

namespace ElectricityCost.CLI.Commands.Core;

public class TemplateManager
{
    private readonly string _filePath;
    private const string CURRENCY_SYMBOL = "£";

    public TemplateManager(string filePath)
    {
        _filePath = filePath;
    }
    
    public void CreateExceltemplate()
    {
        using (var workbook = new XLWorkbook())
        {
            int numberOfDaysInCurrentMonth = DateTime.DaysInMonth(DateTime.Today.Year, DateTime.Today.Month);
            var firstDayOfMonth = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            var worksheet = workbook.Worksheets.Add(DateTime.Today.ToString("MMMM yyyy"));
            worksheet.Author = "DevPower.co.uk";
    
            var rngTable = worksheet.Range("A1:E100");
    
            worksheet.Cell("A1").Value = "Date";
    
            var rngHeaders = rngTable.Range("A1:B1");
            rngHeaders.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rngHeaders.Style.Fill.BackgroundColor = XLColor.Blue;
            rngHeaders.Style.Font.Bold = true;
            rngHeaders.Style.Font.FontSize = 14;
            rngHeaders.Style.Font.FontColor = XLColor.White;

            var rngTotals = rngTable.Range("D1:E1");
            rngTotals.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rngTotals.Style.Font.Bold = true;
            rngTotals.Style.Font.FontSize = 32;
            rngTotals.Style.Font.FontColor = XLColor.DarkRed;
            rngTotals.Style.Border.OutsideBorder = XLBorderStyleValues.Thick;
            rngTotals.Style.NumberFormat.Format = $"{CURRENCY_SYMBOL} #,##0.00";
            
            var rngSubTotals = rngTable.Range("D2:E3");
            rngSubTotals.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            rngSubTotals.Style.Font.Bold = true;
            rngSubTotals.Style.Font.FontSize = 20;
            rngSubTotals.Style.Font.FontColor = XLColor.Red;
            rngSubTotals.Style.NumberFormat.Format = $"{CURRENCY_SYMBOL} #,##0.00";
            
            var dateRange = rngTable.Range($"B2:B{numberOfDaysInCurrentMonth + 1}");
            for (int i = 1; i <= numberOfDaysInCurrentMonth; i++)
            {
                worksheet.Cell($"A{i+1}").Value = firstDayOfMonth.AddDays(i - 1).Date; //.ToString("dd/MM/yyyy");
                worksheet.Cell($"B{i + 1}").Value = 0.00m; // GetRandomCost();
                dateRange.Style.NumberFormat.Format = $"{CURRENCY_SYMBOL} #,##0.00";
            }
    
            worksheet.Cell("B1").Value = "Daily Total Cost";
            worksheet.Cell("D1").Value = "MONTHLY TOTAL COST:";
            worksheet.Cell("E1").FormulaA1 = $"=SUM(B2:B{numberOfDaysInCurrentMonth + 1})";
            worksheet.Cell("D2").Value = "DAILY AVERAGE:";
            worksheet.Cell("E2").FormulaA1 = $"=AVERAGEIF(B2:B{numberOfDaysInCurrentMonth + 1},\"<>0\")";
            worksheet.Cell("D3").Value = "EXPECTED TOTAL:";
            worksheet.Cell("E3").FormulaA1 = $"=E2*{numberOfDaysInCurrentMonth}";
            
            worksheet.Column(1).AdjustToContents();
            worksheet.Column(2).AdjustToContents();
            worksheet.Column(3).Width = 10;
            worksheet.Column(4).AdjustToContents();
            worksheet.Column(5).AdjustToContents();
            
            workbook.SaveAs(_filePath);
        }
    }

    public void AddCost(decimal dayOfMonth, decimal cost)
    {
        using (var workbook = new XLWorkbook(_filePath))
        {
            var worksheet = workbook.Worksheets.First(ws => ws.Name == DateTime.Today.ToString("MMMM yyyy"));
            worksheet.Cell($"B{dayOfMonth + 1}").Value = cost;
            workbook.SaveAs(_filePath);
        }
    }
}