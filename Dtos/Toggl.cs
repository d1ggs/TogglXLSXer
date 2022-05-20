namespace Dtos
{
    public class TogglWorkspaceDto
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public int Profile { get; set; }
        public bool Premium { get; set; }
        public bool Admin { get; set; }
        public int DefaultHourlyRate { get; set; }
        public string DefaultCurrency { get; set; }
        public bool OnlyAdminsMayCreateProjects;
        public bool OnlyAdminsSeeBillableRates;
        public bool OnlyAdminsSeeTeamDashBoard;
        public bool ProjectsBillableByDefault;
        public int Rounding;
        public int RoundingMinutes;
        public string At;
        public bool IcalEnabled;
        public string ApiToken;

    }
}