namespace Excel2SQLExporter.Models;

// ─── Voucher Mode ─────────────────────────────────────────────────────────────

/// <summary>
/// Controls how Journal Voucher records are written to the database.
///
/// PerProduct  — one JUR number per Excel row.
///               BillNumbers is incremented once per row.
///               Each row gets its own Voucher record and 2 AccountsTransaction rows.
///               Narration: "Stock Opening Balance - Product Code: RM55597167996"
///
/// BatchSingle — one JUR number shared by the entire import batch.
///               BillNumbers is incremented exactly ONCE before the loop.
///               Every StockTransaction row shares the same voucher number.
///               ONE Voucher record written at the end (total batch value).
///               TWO AccountsTransaction rows written at the end (total batch value).
///               Narration: "Stock Opening Balance - Batch Import dd/MM/yyyy (N products)"
/// </summary>
public enum VoucherMode
{
    PerProduct,   // default — one voucher per row
    BatchSingle   // one voucher for the entire import
}

// ─── Import Status ────────────────────────────────────────────────────────────

public enum ImportStatus
{
    NewProduct,       // Product + Stock + StockTransaction (+ Voucher/AccountsTx if PerProduct)
    ExistingProduct,  // Stock + StockTransaction only (+ Voucher/AccountsTx if PerProduct)
    Skipped,          // Invalid row — nothing written
    Error             // Exception — transaction rolled back
}

// ─── Import Result ────────────────────────────────────────────────────────────

public class ImportResult
{
    public int          RowNumber   { get; set; }
    public string       ProductCode { get; set; } = string.Empty;
    public string       ProductName { get; set; } = string.Empty;
    public ImportStatus Status      { get; set; }
    public string       Message     { get; set; } = string.Empty;
}

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

// ─── Import Summary ───────────────────────────────────────────────────────────

public class ImportSummary
{
    public int        TotalRows             { get; set; }
    public int        NewProducts           { get; set; }
    public int        ExistingProducts      { get; set; }
    public int        Skipped               { get; set; }
    public int        Errors                { get; set; }
    public VoucherMode VoucherMode          { get; set; }
    public string     BatchVoucherNo        { get; set; } = string.Empty;
    public decimal    BatchTotalValue       { get; set; }

    // Barcode stats — only populated when barcode insert is enabled
    public bool       BarcodesEnabled       { get; set; }
    public int        TotalBarcodesInserted { get; set; }   // sum of all Qty iterations

    public List<ImportResult> Results   { get; set; } = [];
    public DateTime           StartTime { get; set; }
    public DateTime           EndTime   { get; set; }

    public TimeSpan Duration => EndTime - StartTime;
}
