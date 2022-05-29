using System.Data;
using System.Globalization;
using OfficeOpenXml;

namespace Lib;

public class ReportFormatter
{
    
    private readonly int _pageHeight = 34;
    private readonly char _pageWidth = 'L';
    private readonly int[] _excelColumns = Enumerable.Range(2, 9).ToArray();
    private readonly string _company;
    private readonly string _person;

    private readonly string[] _dataColumns = new[]
        { "DATA", "CLIENTE", "PROGETTO", "DESCRIZIONE", "H. INIZIO", "H. FINE", "FERIE/PERMESSI", "IN PRESENZA S/N" };

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

    public string MonthName =>
        new DateOnly(1, _month, 2000).ToString("MMMM", _italianCultureInfo);
    
    public ReportFormatter(string company, string person, int month, int year)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        _person = person;
        _company = company;
        _month = month;
        _year = year;
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
    
    public void FormatCsvToExcel(DataTable dataTable)
    {

        using(var package = new ExcelPackage(new FileInfo("TimeReport.xlsx")))
        {
            var workbook = package.Workbook;
            var worksheet = workbook.Worksheets.Add("Foglio 1");
            
            AddSheetHeader(worksheet);

            var rowCounter = 9;

            foreach (DataRow row in dataTable.Rows)
            {
                foreach (var (dataColumn, excelColumn) in _dataColumns.Zip(_excelColumns))
                {

                    var tableColumn = _columnMapping[dataColumn];
                    var value = row.Field<string>(tableColumn);
                    
                    if (dataColumn == "DATA")
                    {
                        value = DateOnly.ParseExact(value, "yyyy/MM/dd", _italianCultureInfo)
                            .ToString("dd/MM/yyyy");
                    }
                    
                    // TODO Round start and end time
                    // 
                    
                    worksheet.SetValue(rowCounter, excelColumn, value);
                    
                    rowCounter++;
                }
            }
        }
    }
}