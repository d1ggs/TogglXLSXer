using System.Net.Http.Headers;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using Lib.Dtos;
namespace Lib;

public class EmptyReportException : Exception
{
    
}

public class CsvReportDownloader
{
    private const string ReportEndpoint = "https://api.track.toggl.com/reports/api/v2/details.csv";
    private const string WorkspaceEndpoint = "https://api.track.toggl.com/api/v8/workspaces";

    private readonly string _apiKey;
    private readonly HttpClient _client;
    private readonly bool _debug;

    public CsvReportDownloader(string apiKey, bool debug=false)
    {
        _apiKey = apiKey;
        _client = new HttpClient();
        _debug = debug;
    }

    public async Task<List<TogglWorkspaceDto>> GetWorkspaces()
    {
        // Add authorization
        var byteArray = Encoding.ASCII.GetBytes($"{_apiKey}:api_token");
        var base64Auth = Convert.ToBase64String(byteArray);
        _client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", base64Auth);
        
        var workspacesTask = _client.GetStringAsync(WorkspaceEndpoint);
        var response = await workspacesTask;
        
        // Since the API responds with snake case JSON, we need to deserialize to Pascal case DTOs
        var settings = new JsonSerializerSettings
        {
            ContractResolver = new DefaultContractResolver { NamingStrategy = new SnakeCaseNamingStrategy() } 
        };
        
        // Deserialize the list of workspaces on the corresponding DTOs
        var dtos = JsonConvert.DeserializeObject<List<TogglWorkspaceDto>>(response, settings);

        return dtos;
    }
    
    public async Task<string> DownloadDetailedReport(string workspaceId, int year, int month)
    {
        var byteArray = Encoding.ASCII.GetBytes($"{_apiKey}:api_token");
        var base64Auth = Convert.ToBase64String(byteArray);
        
        // Build the date range strings needed to download the report
        var startDate = new DateTime(year, month, 1).ToString("yyyy-MM-dd");
        var endDate = new DateTime(year, month, DateTime.DaysInMonth(year, month)).ToString("yyyy-MM-dd");
        
        _client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", base64Auth);
        // Required parameters are passed in the GET request and are user_agent and workspace_id
        var reportTask = _client.GetStringAsync($"{ReportEndpoint}?" +
                                                $"user_agent=stornello-ducati0a@icloud.com&" +
                                                $"workspace_id={workspaceId}&" +
                                                $"since={startDate}&" +
                                                $"until={endDate}");
        
        var report = await reportTask;
        
        // Check that the number of rows is at least 3: we need the header plus at least an entry.
        // There is also a blank line at the end of the report, so we need to keep this into account.
        var rowsNumber = report.Split('\n').Length;
        if (_debug) Console.WriteLine($"Report has length {rowsNumber}");
        if (rowsNumber < 3) throw new EmptyReportException();

        return report;
    }
    
}

