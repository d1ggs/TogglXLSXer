using System.Net.Http.Headers;
using System.Text;
using Dtos;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
namespace Lib;

public class CsvReportDownloader
{
    public const string ReportEndpoint = "https://api.track.toggl.com/reports/api/v2/details.csv";
    public const string WorkspaceEndpoint = "https://api.track.toggl.com/api/v8/workspaces";
    
    private readonly string _apiKey;
    private readonly HttpClient _client;
    
    public CsvReportDownloader(string apiKey)
    {
        _apiKey = apiKey;
        _client = new HttpClient();
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
    
    public async Task<string> DownloadDetailedReport(string workspaceId)
    {
        var byteArray = Encoding.ASCII.GetBytes($"{_apiKey}:api_token");
        var base64Auth = Convert.ToBase64String(byteArray);
        // Required parameters are passed in the GET request and are user_agent and workspace_id
        _client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", base64Auth);
        var reportTask = _client.GetStringAsync($"{ReportEndpoint}?user_agent=diego.piccinotti@gmail.com&" +
                                                 $"workspace_id={workspaceId}");
        
        var report = await reportTask;

        return report;
    }
    
}

