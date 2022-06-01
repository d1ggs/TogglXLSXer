namespace Lib.Columns;

internal static class ExpenseColumns
{
    internal const string Date = "DATA";
    internal const string Project = "PROGETTO";
    internal const string Location = "LUOGO";
    internal const string Description = "DESCRIZIONE SPESA";
    internal const string Amount = "EURO";

    internal static readonly string[] All = new[] {
        Date,
        Project,
        Location,
        Description,
        Amount
    };

    internal static readonly int[] Indices = Enumerable.Range(2, 5).ToArray();
}
