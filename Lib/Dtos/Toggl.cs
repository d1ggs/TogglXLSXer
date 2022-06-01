namespace Lib.Dtos;

public class TogglWorkspaceDto
{
    public int Id { get; set; }
    public string? Name { get; set; }
    public int Profile { get; set; }
    public bool Premium { get; set; }
    public bool Admin { get; set; }
    public int DefaultHourlyRate { get; set; }
    public string? DefaultCurrency { get; set; }
    public bool OnlyAdminsMayCreateProjects { get; set; }
    public bool OnlyAdminsSeeBillableRates { get; set; }
    public bool OnlyAdminsSeeTeamDashBoard { get; set; }
    public bool ProjectsBillableByDefault { get; set; }
    public int Rounding { get; set; }
    public int RoundingMinutes { get; set; }
    public string? At { get; set; }
    public bool IcalEnabled { get; set; }
    public string? ApiToken { get; set; }
}
