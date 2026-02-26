using Microsoft.Data.SqlClient;

namespace Excel2SQLExporter.Services;

/// <summary>
/// Handles destructive DB operations — Recall and Delete All.
///
/// Recall  — deletes every DB record that was written for a given set of
///            ProductCodes.  Lookup tables (Brands/Categories/Groups) are left
///            intact because they may be shared with other products outside
///            this import.
///
/// Delete All — TRUNCATEs every table the exporter ever writes to, including
///              BillNumbers (the importer will recreate it starting at
///              JUR0000001 on the next run).
///
/// Deletion order (child → parent, safest even without FK constraints):
///   1. PrintBarCodes          — barcode label queue
///   2. AccountsTransaction    — double-entry rows
///   3. Voucher                — journal voucher header
///   4. StockTransaction       — stock movement audit
///   5. Stock                  — stock balances
///   6. Product                — product master
///   7. ProductBrands  \
///   8. ProductCategories  }   — Delete All only
///   9. ProductGroups  /
///  10. BillNumbers            — Delete All only (counter row)
/// </summary>
public class SqlDeleteService(string connectionString)
{
    // ─── Recall: delete by matching ProductCodes ──────────────────────────────

    public async Task<DeleteSummary> RecallByProductCodesAsync(
        List<string>                                        productCodes,
        IProgress<(int Current, int Total, string Message)> progress,
        CancellationToken                                   ct = default)
    {
        var summary = new DeleteSummary { StartTime = DateTime.Now, Mode = DeleteMode.Recall };

        if (productCodes.Count == 0)
        {
            summary.EndTime = DateTime.Now;
            return summary;
        }

        await using var conn = new SqlConnection(connectionString);
        await conn.OpenAsync(ct);

        // Load all codes into a temp table — avoids an IN list of 500+ parameters
        progress.Report((0, 7, "📋  Building product code list..."));
        await CreateTempCodesAsync(conn, productCodes, ct);

        await using var tx = conn.BeginTransaction();
        try
        {
            // Step 1 — Collect voucher numbers tied to these products BEFORE deleting
            progress.Report((1, 7, "🔍  Collecting voucher numbers..."));
            var voucherNumbers = await CollectVoucherNumbersAsync(conn, tx, ct);

            // Step 2 — PrintBarCodes
            progress.Report((2, 7, "🗑️  Deleting PrintBarCodes rows..."));
            summary.DeletedBarcodes = await ExecDeleteAsync(conn, tx,
                "DELETE [dbo].[PrintBarCodes] WHERE [BarCodeString] IN (SELECT [ProductCode] FROM #TempCodes)", ct);

            // Step 3 — AccountsTransaction (by voucher number)
            progress.Report((3, 7, "🗑️  Deleting AccountsTransaction rows..."));
            summary.DeletedAccountsTransactions = voucherNumbers.Count > 0
                ? await DeleteByVoucherAsync(conn, tx, "AccountsTransaction", "VoucherNumber", voucherNumbers, ct)
                : 0;

            // Step 4 — Voucher
            progress.Report((4, 7, "🗑️  Deleting Voucher rows..."));
            summary.DeletedVouchers = voucherNumbers.Count > 0
                ? await DeleteByVoucherAsync(conn, tx, "Voucher", "VoucherNumber", voucherNumbers, ct)
                : 0;

            // Step 5 — StockTransaction
            progress.Report((5, 7, "🗑️  Deleting StockTransaction rows..."));
            summary.DeletedStockTransactions = await ExecDeleteAsync(conn, tx,
                "DELETE [dbo].[StockTransaction] WHERE [ProductCode] IN (SELECT [ProductCode] FROM #TempCodes)", ct);

            // Step 6 — Stock
            progress.Report((6, 7, "🗑️  Deleting Stock rows..."));
            summary.DeletedStockRecords = await ExecDeleteAsync(conn, tx,
                "DELETE [dbo].[Stock] WHERE [ProductCode] IN (SELECT [ProductCode] FROM #TempCodes)", ct);

            // Step 7 — Product
            progress.Report((7, 7, "🗑️  Deleting Product rows..."));
            summary.DeletedProducts = await ExecDeleteAsync(conn, tx,
                "DELETE [dbo].[Product] WHERE [ProductCode] IN (SELECT [ProductCode] FROM #TempCodes)", ct);

            await tx.CommitAsync(ct);
        }
        catch
        {
            await tx.RollbackAsync(ct);
            throw;
        }

        summary.EndTime = DateTime.Now;
        return summary;
    }

    // ─── Delete All: TRUNCATE every exporter table ────────────────────────────

    public async Task<DeleteSummary> DeleteAllAsync(
        IProgress<(int Current, int Total, string Message)> progress,
        CancellationToken                                   ct = default)
    {
        var summary = new DeleteSummary { StartTime = DateTime.Now, Mode = DeleteMode.DeleteAll };

        await using var conn = new SqlConnection(connectionString);
        await conn.OpenAsync(ct);

        // TRUNCATE order: children first, then parents.
        // No FK constraints exist so this is for logical clarity.
        var steps = new (string Label, string Sql)[]
        {
            ("PrintBarCodes",        "TRUNCATE TABLE [dbo].[PrintBarCodes]"),
            ("AccountsTransaction",  "TRUNCATE TABLE [dbo].[AccountsTransaction]"),
            ("Voucher",              "TRUNCATE TABLE [dbo].[Voucher]"),
            ("StockTransaction",     "TRUNCATE TABLE [dbo].[StockTransaction]"),
            ("Stock",                "TRUNCATE TABLE [dbo].[Stock]"),
            ("Product",              "TRUNCATE TABLE [dbo].[Product]"),
            ("ProductBrands",        "TRUNCATE TABLE [dbo].[ProductBrands]"),
            ("ProductCategories",    "TRUNCATE TABLE [dbo].[ProductCategories]"),
            ("ProductGroups",        "TRUNCATE TABLE [dbo].[ProductGroups]"),
            // BillNumbers: TRUNCATE removes the counter row; the next import recreates it
            // starting at JUR0000001 — which is correct after a full wipe.
            ("BillNumbers",          "TRUNCATE TABLE [dbo].[BillNumbers]"),
        };

        int total = steps.Length;

        for (int i = 0; i < total; i++)
        {
            ct.ThrowIfCancellationRequested();
            var (label, sql) = steps[i];
            progress.Report((i + 1, total, $"🗑️  Truncating {label}..."));

            await using var cmd = new SqlCommand(sql, conn);
            await cmd.ExecuteNonQueryAsync(ct);
            summary.TruncatedTables.Add(label);
        }

        summary.EndTime = DateTime.Now;
        return summary;
    }

    // ─── Private helpers ──────────────────────────────────────────────────────

    /// Insert all product codes into a session-scoped temp table for fast joins.
    private static async Task CreateTempCodesAsync(
        SqlConnection conn, List<string> codes, CancellationToken ct)
    {
        await using var create = new SqlCommand(
            "CREATE TABLE #TempCodes ([ProductCode] NVARCHAR(15))", conn);
        await create.ExecuteNonQueryAsync(ct);

        // Batch-insert in groups of 100 to avoid parameter limits
        const int batchSize = 100;
        for (int offset = 0; offset < codes.Count; offset += batchSize)
        {
            var batch = codes.Skip(offset).Take(batchSize).ToList();
            var values = string.Join(",", batch.Select((_, j) => $"(@p{j})"));
            var sql    = $"INSERT INTO #TempCodes ([ProductCode]) VALUES {values}";

            await using var ins = new SqlCommand(sql, conn);
            for (int j = 0; j < batch.Count; j++)
                ins.Parameters.AddWithValue($"@p{j}", batch[j]);
            await ins.ExecuteNonQueryAsync(ct);
        }
    }

    /// Collect all distinct VoucherNumbers from StockTransaction rows
    /// that belong to the currently-loaded product codes.
    private static async Task<List<string>> CollectVoucherNumbersAsync(
        SqlConnection conn, SqlTransaction tx, CancellationToken ct)
    {
        var list = new List<string>();
        await using var cmd = new SqlCommand(
            """
            SELECT DISTINCT [VoucherNumber]
            FROM  [dbo].[StockTransaction]
            WHERE [VoucherNumber] IS NOT NULL
              AND [ProductCode] IN (SELECT [ProductCode] FROM #TempCodes)
            """, conn, tx);

        await using var reader = await cmd.ExecuteReaderAsync(ct);
        while (await reader.ReadAsync(ct))
        {
            string? v = reader[0]?.ToString();
            if (!string.IsNullOrWhiteSpace(v)) list.Add(v);
        }
        return list;
    }

    /// DELETE rows from a table where VoucherNumber is in the provided list.
    /// Uses a temp table for the voucher numbers to avoid a huge IN list.
    private static async Task<int> DeleteByVoucherAsync(
        SqlConnection conn, SqlTransaction tx,
        string table, string column,
        List<string> voucherNumbers, CancellationToken ct)
    {
        // Build a VALUES list directly — voucher numbers are short strings, bounded count
        var paramNames = voucherNumbers.Select((_, i) => $"@v{i}").ToList();
        string sql = $"DELETE [dbo].[{table}] WHERE [{column}] IN ({string.Join(",", paramNames)})";

        await using var cmd = new SqlCommand(sql, conn, tx);
        for (int i = 0; i < voucherNumbers.Count; i++)
            cmd.Parameters.AddWithValue($"@v{i}", voucherNumbers[i]);

        return await cmd.ExecuteNonQueryAsync(ct);
    }

    private static async Task<int> ExecDeleteAsync(
        SqlConnection conn, SqlTransaction tx, string sql, CancellationToken ct)
    {
        await using var cmd = new SqlCommand(sql, conn, tx);
        return await cmd.ExecuteNonQueryAsync(ct);
    }
}

// ─── Delete Summary model ─────────────────────────────────────────────────────

public enum DeleteMode { Recall, DeleteAll }

public class DeleteSummary
{
    public DeleteMode Mode                    { get; set; }
    public int        DeletedProducts         { get; set; }
    public int        DeletedStockRecords     { get; set; }
    public int        DeletedStockTransactions{ get; set; }
    public int        DeletedVouchers         { get; set; }
    public int        DeletedAccountsTransactions { get; set; }
    public int        DeletedBarcodes         { get; set; }
    public List<string> TruncatedTables       { get; set; } = [];
    public DateTime   StartTime               { get; set; }
    public DateTime   EndTime                 { get; set; }
    public TimeSpan   Duration                => EndTime - StartTime;

    public string ToLogSummary() => Mode == DeleteMode.DeleteAll
        ? $"Truncated {TruncatedTables.Count} tables: {string.Join(", ", TruncatedTables)}"
        : $"Products: {DeletedProducts}  Stock: {DeletedStockRecords}  " +
          $"StockTx: {DeletedStockTransactions}  Vouchers: {DeletedVouchers}  " +
          $"AcctTx: {DeletedAccountsTransactions}  Barcodes: {DeletedBarcodes}";
}
