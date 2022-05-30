using System.Data;

namespace Lib;

public static class ReportConverter
{
    public static DataTable BuildDataTableFromCsv(string csv)
    {

        var table = new DataTable("report");
        
        var rows = csv.Split('\n');

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

        table.DefaultView.Sort = "Start date,Start time";

        return table.DefaultView.ToTable();
    }

}