using Lib;

// See https://aka.ms/new-console-template for more information

namespace Application;

class Program
{
    
    static async Task Main(string[] args)
    {
        
        Console.WriteLine("Insert your API Token.\n" +
                          "You can find it at the bottom of the " +
                          "https://track.toggl.com/profile page");
        var apiKey = Console.ReadLine();
        if (apiKey == null) throw new ArgumentNullException();
        Console.WriteLine();
        
        var reportDownloader = new CsvReportDownloader(apiKey);
        
        // Get the workspaces and select the desired one
        var workspaces = await reportDownloader.GetWorkspaces();
        
        Console.WriteLine("Available workspaces:");
        for (int i = 0; i < workspaces.Count; i++)
        {
            var workspace = workspaces[i];
            Console.WriteLine($"{i} - {workspace.Name}");
        }
        
        Console.WriteLine();
        Console.WriteLine("Insert the number of the workspace you desire");
        
        var workspaceIndexStr = Console.ReadLine();
        
        // Do not allow null values
        if (workspaceIndexStr == null) throw new ArgumentNullException();
        var workspaceIndex = Int32.Parse(workspaceIndexStr);
        
        var workspaceId = workspaces[workspaceIndex].Id.ToString();
        
        var report = await reportDownloader.DownloadDetailedReport(workspaceId);
        
        // await File.WriteAllTextAsync("report.csv", report);
        
        var dataTable = ReportConverter.BuildDataTableFromCsv(report);
        ReportConverter.ShowData(dataTable);
        
        return;
    }
}