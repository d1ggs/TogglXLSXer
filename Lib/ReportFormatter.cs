using System.Data;
using System.Globalization;
using System.Net;
using OfficeOpenXml;

namespace Lib;

public class ReportFormatter
{
    
    private readonly int _pageHeight = 34;
    private readonly char _pageWidth = 'L';
    private readonly int[] _excelColumns = Enumerable.Range(2, 10).ToArray();
    private readonly string _company;
    private readonly string _person;

    private readonly string[] _dataColumns = new[]
        { "DATA", "CLIENTE", "PROGETTO", "DESCRIZIONE", "H. INIZIO", "H. FINE", "TOTALE", "FERIE/PERMESSI", "IN PRESENZA S/N" };

    private readonly Dictionary<string, string> _columnMapping = new Dictionary<string, string>
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

    public string MonthName =>
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
        worksheet.SetValue(4, 2, "Società"); // bold_format
        worksheet.SetValue(4, 3, _company); // light_green_regular_format
        worksheet.SetValue(5, 2, "Risorsa"); //, bold_format)
        worksheet.SetValue(5, 3, _person); // light_green_regular_format)
        worksheet.SetValue(6, 2, "Mese"); //, bold_format)
        worksheet.SetValue(6, 3, $"{MonthName} {_year}"); // light_green_regular_format
    }

    private void AddTableHeader(ExcelWorksheet worksheet, int rowIndex)
    {
        foreach (var (columnName, columnIndex) in _dataColumns.Zip(_excelColumns)) 
            worksheet.SetValue(rowIndex, columnIndex, columnName);
    }
    
    public Task FormatCsvToExcel(DataTable dataTable)
    {
        
        using(var package = new ExcelPackage(new FileInfo("TimeReport.xlsx")))
        {
            var workbook = package.Workbook;
            var worksheet = workbook.Worksheets.Add("Foglio 1");
            
            AddSheetHeader(worksheet);

            var rowCounter = 9;
            
            AddTableHeader(worksheet, rowCounter);

            rowCounter++;

            foreach (DataRow row in dataTable.Rows)
            {
                WriteSheetRow(row, worksheet, rowCounter);
                rowCounter++;
            }

            return package.SaveAsAsync("timesheet.xlsx");
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

            switch (dataColumn)
            {
                case "DATA":
                    value = DateOnly.ParseExact(
                        value,
                        "yyyy-MM-dd", _italianCultureInfo).ToString("dd/MM/yyyy");
                    break;

                // Round to nearest minute the time
                case "H. INIZIO":
                    value = RoundTimeOnlyString(value);
                    StartTime = TimeOnly.ParseExact(value, "HH:mm", _italianCultureInfo);
                    break;
                
                case "H. FINE":
                    value = RoundTimeOnlyString(value);
                    EndTime = TimeOnly.ParseExact(value, "HH:mm", _italianCultureInfo);
                    break;
                
                case "TOTALE":
                    tags = row.Field<string>("Tags") ?? "";

                    // Skip total column if vacation or office leave was used
                    if (tags.Contains("ferie") || tags.Contains("permesso")) value = "";
                    else value = (EndTime - StartTime).ToString("hh\\:mm");
                    break;
                
                case "FERIE/PERMESSI":
                    tags = row.Field<string>("Tags") ?? "";

                    // Skip vacation column if no vacation or office leave was used
                    if (!(tags.Contains("ferie") || tags.Contains("permesso"))) continue;
                    
                    value = (EndTime - StartTime).ToString("hh\\:mm");
                    break;
                
                case "IN PRESENZA S/N":
                    tags = row.Field<string>("Tags") ?? "";

                    // Search for a "remot*" substring in tags to identify remote working days
                    value = (!tags.Contains("remot"))? "S" : "N";
                    break;
                
                default:
                    break;
            }

            worksheet.SetValue(rowCounter, excelColumn, value);
        }
    }

    private string RoundTimeOnlyString(string value)
    {
        var time = TimeOnly.ParseExact(value, "HH:mm:ss", _italianCultureInfo);

        // Leave 23:59 alone, as it would turn to 00:00 but date would not be increased
        if (time.Hour == 23 && time.Minute == 59) return value;

        if (time.Second > 30)
        {
            time = new TimeOnly(time.Hour, time.Minute).AddMinutes(1);
        }

        value = time.ToString("hh\\:mm");
        return value;
    }
}