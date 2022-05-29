using Lib;

// See https://aka.ms/new-console-template for more information

namespace Application;

internal static class Program
{
    private static async Task Main(string[] args)
    {
        
        Console.WriteLine("Insert your API Token.\n" +
                          "You can find it at the bottom of the " +
                          "https://track.toggl.com/profile page");
        var apiKey = Console.ReadLine();
        if (apiKey == null) throw new ArgumentNullException
        {
            HelpLink = null,
            HResult = 0,
            Source = null
        };
        Console.WriteLine();
        
        var reportDownloader = new CsvReportDownloader(apiKey);
        
        // Get the workspaces and select the desired one
        var workspaces = await reportDownloader.GetWorkspaces();
        
        Console.WriteLine("Available workspaces:");
        for (var i = 0; i < workspaces.Count; i++)
        {
            var workspace = workspaces[i];
            Console.WriteLine($"{i} - {workspace.Name}");
        }
        
        Console.WriteLine("\nInsert the number of the workspace you desire");
        
        var workspaceIndexStr = Console.ReadLine();
        
        // Do not allow null values
        if (workspaceIndexStr == null) throw new ArgumentNullException();
        var workspaceIndex = int.Parse(workspaceIndexStr);
        
        var workspaceId = workspaces[workspaceIndex].Id.ToString();

        var parsed = false;
        var monthNumber = DateTime.Now.Month;
        var yearNumber = DateTime.Now.Year;
        
        Console.WriteLine("\nInsert the month number for which you desire to download the report");

        while (!parsed)
        {
            var month = Console.ReadLine();
            parsed = int.TryParse(month, out monthNumber);
            if (!parsed)
            {
                Console.WriteLine("\nYou input an invalid value. Please input a number between 1 and 12");
            }
        }

        parsed = false;
        
        Console.WriteLine("\nInsert the year for which you desire to download the report");

        while (!parsed)
        {
            var year = Console.ReadLine();
            parsed = int.TryParse(year, out yearNumber);
            if (!parsed)
            {
                Console.WriteLine($"\nYou input an invalid value. " +
                                  $"Please input an year between 1970 and {DateTime.Now.Year}");
            }
        }

        string report;
        
        // The invocation of DownloadDetailedReport might fail if the report is empty
        try
        {
            report = await reportDownloader.DownloadDetailedReport(workspaceId, yearNumber, monthNumber);
        }
        catch (EmptyReportException)
        {
            Console.WriteLine("There is no data for the selected date range.");
            return;
        }
        
        // Write the report to a file
        await File.WriteAllTextAsync("report.csv", report);
        
        // Put the report inside a DataTable, for easier handling
        var dataTable = ReportConverter.BuildDataTableFromCsv(report);
    }
}