namespace Lib.Columns;

internal static class DataColumns
{
    internal const string Date = "DATA";
    internal const string Client = "CLIENTE";
    internal const string Project = "PROGETTO";
    internal const string Description = "DESCRIZIONE";
    internal const string StartHour = "H. INIZIO";
    internal const string EndHour = "H. FINE";
    internal const string Total = "TOTALE";
    internal const string Vacation = "FERIE/PERMESSI";
    internal const string OnPremise = "IN PRESENZA";

    internal static readonly string[] All = new[] {
        Date,
        Client,
        Project,
        Description,
        StartHour,
        EndHour,
        Total,
        Vacation,
        OnPremise
    };

    internal static readonly int[] Indices = Enumerable.Range(2, 9).ToArray();
}
