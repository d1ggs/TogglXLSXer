using System.Data;
using Excel = Microsoft.Office.Interop.Excel;

namespace Lib;

public class ReportConverter
{
    public static DataTable BuildDataTableFromCsv(string csv)
    {
        
        var table = new DataTable("report");

        var rows = csv.Split('\n');
        Console.WriteLine($"Number of entries in the report: {rows.Length}");
        
        // Get the header, then remove it from the rows
        var columns = rows[0].Split(',');
        
        foreach (var column in columns)
        {
            table.Columns.Add(column);
        }
        
        // Add each row in the datatable
        for (int i = 1; i < rows.Length; i++)
        {
            var row = rows[i];
            var values = row.Split(',');
            var dataTableRow = table.NewRow();
            dataTableRow.ItemArray = values;
            table.Rows.Add(dataTableRow);
        }

        return table;
    }
    
    public static void writeToExcel()
    {
            
        Excel.Application myexcelApplication = new Excel.Application();
        if (myexcelApplication != null)
        {
            Excel.Workbook myexcelWorkbook = myexcelApplication.Workbooks.Add();
            Excel.Worksheet myexcelWorksheet = (Excel.Worksheet)myexcelWorkbook.Sheets.Add();

            myexcelWorksheet.Cells[1, 1] = "Value 1";
            myexcelWorksheet.Cells[2, 1] = "Value 2";
            myexcelWorksheet.Cells[3, 1] = "Value 3";

            myexcelApplication.ActiveWorkbook.SaveAs(@"C:\abc.xlsx", Excel.XlFileFormat.xlWorkbookDefault);

            myexcelWorkbook.Close();
            myexcelApplication.Quit();
        }
    }

    public static void ShowData(DataTable dtData)
    {
        if (dtData.Rows.Count <= 0) return;
        foreach (DataColumn dc in dtData.Columns)
        {
            Console.Write(dc.ColumnName + " ");
        }
        Console.WriteLine("\n-----------------------------------------------");

        foreach (DataRow dr in dtData.Rows)
        {
            foreach (var item in dr.ItemArray)
            {
                Console.Write(item.ToString() + "      ");
            }
            Console.Write("\n");
        }
    }

}