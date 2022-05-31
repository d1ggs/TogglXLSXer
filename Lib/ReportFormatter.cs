using System.Data;
using System.Drawing;
using System.Globalization;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Lib;

public class ReportFormatter
{
    private const int PageHeight = 33;
    private const int PageWidth = 12;
    private readonly int[] _excelColumns = Enumerable.Range(2, 9).ToArray();
    private readonly string _company;
    private readonly string _person;

    private readonly int[] _columnWidths = { 0, 11, 13, 14, 30, 9, 8, 8, 15, 12, 0 };

    private readonly string[] _dataColumns =
        { "DATA", "CLIENTE", "PROGETTO", "DESCRIZIONE", "H. INIZIO", "H. FINE", "TOTALE", "FERIE/PERMESSI", "IN PRESENZA" };

    private readonly string[] _expenseColumnsNames = { "DATA", "PROGETTO", "LUOGO", "DESCRIZIONE SPESA", "EURO" };
    private readonly int[] _expenseColumns = { 2, 3, 4, 5, 6 };
    
    private readonly Dictionary<string, string> _columnMapping = new()
    {
        { "CLIENTE", "Client" },
        { "PROGETTO", "Project" },
        { "DATA", "Start date"},
        { "H. INIZIO", "Start time" },
        { "H. FINE", "End time" },
        {"DESCRIZIONE", "Description"}
    };
    
    private readonly int _month;
    private readonly int _year;
    private readonly CultureInfo _italianCultureInfo = new CultureInfo("it-IT");
    private readonly bool _debug;
    private TimeSpan _totalWorkedTime;
    
    private readonly Color _darkGreen = ColorTranslator.FromHtml("#548235");
    private readonly Color _lightGreen = ColorTranslator.FromHtml("#92D050");
    private readonly Color _paleGreen = ColorTranslator.FromHtml("#C6E0B4");
    private readonly Color _lightGrey = ColorTranslator.FromHtml("#D9D9D9");

    private string MonthName =>
        new DateOnly(2000, _month, 1).ToString("MMMM", _italianCultureInfo);
    
    public ReportFormatter(string company, string person, int month, int year, bool debug=false)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        _person = person;
        _company = company;
        _month = month;
        _year = year;
        _debug = false;
    }

    private void AddSheetHeader(ExcelWorksheet worksheet)
    {
        worksheet.SetValue(4, 2, "Società");
        worksheet.SetValue(4, 3, _company);
        worksheet.SetValue(5, 2, "Risorsa");
        worksheet.SetValue(5, 3, _person);
        worksheet.SetValue(6, 2, "Mese");
        worksheet.SetValue(6, 3, $"{MonthName} {_year}");
        
        // Add borders to all cells
        worksheet.Cells[4, 2, 6, 3].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        worksheet.Cells[4, 2, 6, 3].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        worksheet.Cells[4, 2, 6, 3].Style.Border.Left.Style = ExcelBorderStyle.Thin;
        worksheet.Cells[4, 2, 6, 3].Style.Border.Right.Style = ExcelBorderStyle.Thin;
        
        // Add the light green fill to the value cells
        worksheet.Cells[4, 3, 6, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
        worksheet.Cells[4, 3, 6, 3].Style.Fill.BackgroundColor.SetColor(_paleGreen);
        
        // Add bold to the heading cells
        worksheet.Cells[4, 2, 6, 2].Style.Font.Bold = true;
    }

    private void AddTableHeader(ExcelWorksheet worksheet, int rowIndex)
    {
        Color[] tableHeaderColors =
        {
            _lightGreen, _darkGreen, _darkGreen, _darkGreen, 
            _lightGreen, _lightGreen, _darkGreen, _lightGreen,
            _lightGreen
        };

        // TODO format the table header
        for (int i = 0; i < tableHeaderColors.Length; i++)
        {
            var columnIndex = _excelColumns[i];
            var columnName = _dataColumns[i];
            var backgroundColor = tableHeaderColors[i];
            // Write value
            worksheet.SetValue(rowIndex, columnIndex, columnName);
            
            // Add styling
            var cell = worksheet.Cells[rowIndex, columnIndex];
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(backgroundColor);
            
            if (backgroundColor == _darkGreen)
            {
                cell.Style.Font.Bold = true;
                cell.Style.Font.Color.SetColor(Color.White);
            }
        }
            
    }
    
    public Task FormatCsvToExcel(DataTable dataTable)
    {
        
        using(var package = new ExcelPackage(new FileInfo("TimeReport.xlsx")))
        {
            var workbook = package.Workbook;
            var worksheet = workbook.Worksheets.Add("Foglio 1");

            _totalWorkedTime = new TimeSpan();
            
            // Add the columns to the worksheet and set the width
            for (var i = 0; i < _columnWidths.Length; i++)
            {
                worksheet.Column(i+1).Width = _columnWidths[i ];
            }

            AddSheetTitle(worksheet);

            AddSheetHeader(worksheet);

            var rowCounter = 9;
            
            AddTableHeader(worksheet, rowCounter);

            rowCounter++;

            var startDate = "";
            var mergeCounter = rowCounter;
            DateOnly prevDate;

            string? curDate;
            
            foreach (DataRow row in dataTable.Rows)
            {
                curDate = row.Field<string>("Start date");
                if (curDate == null) continue;
                
                if (startDate == "")
                {
                    startDate = curDate;
                    prevDate =  DateOnly.ParseExact(curDate, "yyyy-MM-dd", _italianCultureInfo);
                }
                
                var merged = false;
                
                if (curDate != startDate)
                {
                    // Merge cells that share the same date
                    MergeCells(worksheet, new [] { "B", "I", "J"}, mergeCounter, rowCounter - 1);

                    mergeCounter = rowCounter;
                    startDate = curDate;
                    merged = true;
                    
                    // Split the pages if we reached maximum page height
                    if ((rowCounter + 1) % PageHeight == 0)
                        mergeCounter = SkipRowsAndAddTableHeader(worksheet, ref rowCounter);

                    // Fill the missing days. E.g., if we are jumping from 04/05 to 04/07 we need to add 04/06
                    var curDateOnly = DateOnly.
                        ParseExact(curDate ?? throw new InvalidOperationException(), "yyyy-MM-dd", _italianCultureInfo);
                    
                    var prevDay = curDateOnly.AddDays(-1);
                    var curCellDate = prevDate.AddDays(1);
                    while (curCellDate <= prevDay)
                    {
                        // Write the missing date
                        worksheet.SetValue(rowCounter, 2, curCellDate.ToString("dd/MM/yyyy", _italianCultureInfo));
                        
                        // Add the border to the empty cells and paint them grey, assuming missing cells are holidays
                        foreach (var columnIndex in _excelColumns)
                        {
                            var cell = worksheet.Cells[rowCounter, columnIndex];
                            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(_lightGrey);
                            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        }
                        
                        mergeCounter = ++rowCounter;
                        curCellDate = curCellDate.AddDays(1);
                        
                        // Split the pages if we reached maximum page height
                        if ((rowCounter + 1) % PageHeight == 0)
                            mergeCounter = SkipRowsAndAddTableHeader(worksheet, ref rowCounter);
                    }
                    
                    // Since curDateOnly will be written to the worksheet, the next missing day might be the day after
                    prevDate = curDateOnly;

                }
                
                // Split the pages if we reached maximum page height
                if ((rowCounter + 1) % PageHeight == 0)
                {
                    // Merge up to previous cell
                    if (!merged)
                        MergeCells(worksheet, new [] { "B", "I", "J"}, mergeCounter, rowCounter-1);

                    mergeCounter = SkipRowsAndAddTableHeader(worksheet, ref rowCounter);
                }
                
                WriteSheetRow(row, worksheet, rowCounter);
                rowCounter++;

            }
            
            // Add the hours grand total
            worksheet.SetValue(rowCounter, 8, $"{(int) _totalWorkedTime.TotalHours}:{_totalWorkedTime.Minutes}");
            var totalCell = worksheet.Cells[rowCounter, 8];
            totalCell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            totalCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            totalCell.Style.Fill.BackgroundColor.SetColor(_darkGreen);
            totalCell.Style.Font.Color.SetColor(Color.White);
            totalCell.Style.Font.Bold = true;
                
            // Add the expense footer
            var pageNumber = rowCounter / PageHeight + 1;
            rowCounter = PageHeight * pageNumber + 2;
            
            AddExpenseFooter(worksheet, rowCounter);
            rowCounter++;
            
            // Add some empty table rows
            for (var i = 0; i < 7; i++)
            {
                rowCounter++;
                foreach (var columnIndex in _expenseColumns)
                {
                    // Create "empty" table cells by adding borders
                    var cell = worksheet.Cells[rowCounter, columnIndex];
                    cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    
                    // Add background only to last column
                    if (columnIndex != _expenseColumns.Max()) continue;
                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(_paleGreen);
                }
            }
            
            rowCounter++;
            
            worksheet.SetValue(rowCounter, 6, 0);
            totalCell = worksheet.Cells[rowCounter, 6];
            totalCell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            totalCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            totalCell.Style.Fill.BackgroundColor.SetColor(_darkGreen);
            totalCell.Style.Font.Color.SetColor(Color.White);
            totalCell.Style.Font.Bold = true;
            
            // Add worksheet formatting
            worksheet.Column(PageWidth).PageBreak = true;
            worksheet.PrinterSettings.PaperSize = ePaperSize.A4;
            worksheet.PrinterSettings.Orientation = eOrientation.Landscape;
            worksheet.PrinterSettings.LeftMargin = Convert.ToDecimal(0.31496062992126);
            worksheet.PrinterSettings.RightMargin = Convert.ToDecimal(0.31496062992126);
            worksheet.PrinterSettings.TopMargin = Convert.ToDecimal(0.748031496062992);
            worksheet.PrinterSettings.BottomMargin = Convert.ToDecimal(0.748031496062992);
            worksheet.PrinterSettings.HorizontalCentered = true;
            worksheet.PrinterSettings.VerticalCentered = true;

            foreach (var i in Enumerable.Range(0, pageNumber - 1))
            {
                var breakIndex = (i + 1) * PageHeight;
                worksheet.Column(breakIndex).PageBreak = true;
            }
            
            // Save the timesheet
            return package.SaveAsAsync("output/timesheet.xlsx");
        }
    }

    private void AddSheetTitle(ExcelWorksheet worksheet)
    {
        // Write the title
        worksheet.SetValue(2, 2, "TIME REPORT MENSILE");

        // Add title styling
        var titleCells = worksheet.Cells["B2:J2"];
        titleCells.Merge = true;
        titleCells.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        titleCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        titleCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
        titleCells.Style.Fill.BackgroundColor.SetColor(_darkGreen);
        titleCells.Style.Font.Color.SetColor(Color.White);
        titleCells.Style.Font.Bold = true;
    }

    private void AddExpenseFooter(ExcelWorksheet worksheet, int rowCounter)
    {
        // Write title
        worksheet.SetValue(rowCounter, 2, "SPESE MENSILI");
        
        // Add title styling
        var titleCells = worksheet.Cells[$"B{rowCounter}:F{rowCounter}"];
        titleCells.Merge = true;
        titleCells.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        titleCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        titleCells.Style.Fill.PatternType = ExcelFillStyle.Solid;
        titleCells.Style.Fill.BackgroundColor.SetColor(_darkGreen);
        titleCells.Style.Font.Color.SetColor(Color.White);
        titleCells.Style.Font.Bold = true;

        // Add blank expense table
        rowCounter += 2;
        
        foreach (var (columnName, columnIndex) in _expenseColumnsNames.Zip(_expenseColumns))
        {
            
            // Set column header
            worksheet.SetValue(rowCounter, columnIndex, columnName);
            
            // Add styling
            var cell = worksheet.Cells[rowCounter, columnIndex];
            if (columnName is "DATA" or "EURO")
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(_lightGreen);
            }
            else
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(_darkGreen);
                cell.Style.Font.Color.SetColor(Color.White);
                cell.Style.Font.Bold = true;
            }
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        }
    }

    private int SkipRowsAndAddTableHeader(ExcelWorksheet worksheet, ref int rowCounter)
    {
        int mergeCounter;
        // Skip cells and add new page table header
        rowCounter += 3;
        AddTableHeader(worksheet, rowCounter);
        rowCounter++;
        mergeCounter = rowCounter;
        return mergeCounter;
    }

    private static void MergeCells(ExcelWorksheet worksheet, string[] columns, int mergeCounter, int rowCounter)
    {
        foreach (var column in columns)
        {
            var mergedCells = worksheet.Cells[$"{column}{mergeCounter}:{column}{rowCounter}"];
            mergedCells.Merge = true;
            mergedCells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            mergedCells.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        }
    }

    private void WriteSheetRow(DataRow row, ExcelWorksheet worksheet, int rowCounter)
    {
        TimeOnly StartTime;
        TimeOnly EndTime;

        string? tags;

        foreach (var (dataColumn, excelColumn) in _dataColumns.Zip(_excelColumns))
        {

            // If the column is not in the mapping it will be filled later, so we use the empty string as a placeholder
            var value = "";
            if (_columnMapping.ContainsKey(dataColumn))
            {
                var tableColumn = _columnMapping[dataColumn];
                value = row.Field<string>(tableColumn);
                
                // If the value is null, there is something wrong in the data
                if (value == null)
                {
                    if (_debug) Console.WriteLine($"Something wrong at row {rowCounter}, column {dataColumn}");
                    continue;
                }
            }

            var applyAlignment = false;
            var addBackground = false;
            
            switch (dataColumn)
            {
                case "DATA":
                    value = DateOnly.ParseExact(
                        value,
                        "yyyy-MM-dd", _italianCultureInfo).ToString("dd/MM/yyyy");
                    applyAlignment = true;
                    break;

                // Round time to the nearest minute
                case "H. INIZIO":
                    value = RoundTimeOnlyString(value);
                    StartTime = TimeOnly.ParseExact(value, "HH:mm", _italianCultureInfo);
                    break;
                
                case "H. FINE":
                    value = RoundTimeOnlyString(value);
                    EndTime = TimeOnly.ParseExact(value, "HH:mm", _italianCultureInfo);
                    break;
                
                // Compute working/vacation time
                case "TOTALE":
                    tags = row.Field<string>("Tags") ?? "";

                    // Skip total column if vacation or office leave was used
                    if (tags.Contains("ferie") || tags.Contains("permesso")) value = "";
                    
                    else
                    {
                        var span = (EndTime - StartTime);
                        _totalWorkedTime += span;
                        value = span.ToString("hh\\:mm");
                    }

                    addBackground = true;
                    break;
                
                case "FERIE/PERMESSI":
                    tags = row.Field<string>("Tags") ?? "";

                    // Skip vacation column if no vacation or office leave was used
                    if (!(tags.Contains("ferie") || tags.Contains("permesso"))) value="";
                    else value = (EndTime - StartTime).ToString("hh\\:mm");
                    applyAlignment = true;
                    break;
                
                // Flag remote working
                case "IN PRESENZA":
                    tags = row.Field<string>("Tags") ?? "";
                    tags = tags.ToLower();

                    // Search for a "remot*" substring in tags to identify remote working days
                    value = (!tags.Contains("remot"))? "S" : "N";
                    applyAlignment = true;
                    break;
                
                default:
                    break;
            }
            worksheet.SetValue(rowCounter, excelColumn, value);
            // Add the border to the cell
            var cell = worksheet.Cells[rowCounter, excelColumn];
            cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
            if (applyAlignment) cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            if (addBackground)
            {
                cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                cell.Style.Fill.BackgroundColor.SetColor(_paleGreen);
            }
        }
    }

    private string RoundTimeOnlyString(string value)
    {
        var time = TimeOnly.ParseExact(value, "HH:mm:ss", _italianCultureInfo);

        // Leave 23:59 alone, as it would turn to 00:00 but date would not be increased
        if (time.Hour == 23 && time.Minute == 59) return "23:59";

        if (time.Second > 30)
        {
            time = new TimeOnly(time.Hour, time.Minute).AddMinutes(1);
        }

        value = time.ToString("HH\\:mm", _italianCultureInfo);
        return value;
    }
}