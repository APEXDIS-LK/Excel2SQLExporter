namespace ExcelToSqlImporter.Models;

public enum ImportStatus
{
    NewProduct,       // Product + Stock + StockTransaction + Voucher + AccountsTransaction
    ExistingProduct,  // Stock + StockTransaction + Voucher + AccountsTransaction only
    Skipped,          // Invalid row — nothing written to DB
    Error             // Exception during import — transaction rolled back
}

public class ImportResult
{
    public int          RowNumber   { get; set; }
    public string       ProductCode { get; set; } = string.Empty;
    public string       ProductName { get; set; } = string.Empty;
    public ImportStatus Status      { get; set; }
    public string       Message     { get; set; } = string.Empty;
}

/// <summary>
/// Display helpers for ImportResult.
/// Using static extension methods — safe, definitely compiles in .NET 10 / C# 14.
/// </summary>
public static class ImportResultExtensions
{
    public static string StatusIcon(this ImportResult r) => r.Status switch
    {
        ImportStatus.NewProduct      => "✅",
        ImportStatus.ExistingProduct => "🔄",
        ImportStatus.Skipped         => "⚠️",
        ImportStatus.Error           => "❌",
        _                            => "?"
    };

    public static string StatusLabel(this ImportResult r) => r.Status switch
    {
        ImportStatus.NewProduct      => "New Product",
        ImportStatus.ExistingProduct => "Stock Added",
        ImportStatus.Skipped         => "Skipped",
        ImportStatus.Error           => "Error",
        _                            => "Unknown"
    };

    public static string Display(this ImportResult r)
        => $"{r.StatusIcon()} Row {r.RowNumber} [{r.ProductCode}] {r.ProductName} — {r.Message}";
}

public class ImportSummary
{
    public int  TotalRows        { get; set; }
    public int  NewProducts      { get; set; }
    public int  ExistingProducts { get; set; }
    public int  Skipped          { get; set; }
    public int  Errors           { get; set; }

    // C# 14: collection expression initialiser
    public List<ImportResult> Results   { get; set; } = [];
    public DateTime           StartTime { get; set; }
    public DateTime           EndTime   { get; set; }

    public TimeSpan Duration => EndTime - StartTime;
}
