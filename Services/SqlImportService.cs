using Microsoft.Data.SqlClient;
using ExcelToSqlImporter.Models;

namespace ExcelToSqlImporter.Services;

/// <summary>
/// Handles all SQL Server import operations for SimplePOSDB.
/// Per Excel row, writes to 9 tables in a single atomic transaction.
///
/// C# 14: Primary constructor — connectionString injected directly.
/// </summary>
public class SqlImportService(string connectionString)
{
    // ── Fixed account codes for Opening Stock Balance journal entries ─────────
    private const string DebitAccountCode  = "141";  // Stock (Asset — increases on debit)
    private const string CreditAccountCode = "92";   // Opening Balance Equity (increases on credit)
    private const string JournalVoucherType = "Opening Stock Balance";
    private const string AccountTxType      = "Opening Stock Balance";
    private const string ImportUserId       = "IMPORT";

    // ─── Test Connection ──────────────────────────────────────────────────────

    public async Task<(bool Success, string Message)> TestConnectionAsync()
    {
        try
        {
            await using var conn = new SqlConnection(connectionString);
            await conn.OpenAsync();

            await using var cmd = new SqlCommand(
                "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Product'",
                conn);

            int tableCount = (int)(await cmd.ExecuteScalarAsync())!;

            return tableCount == 0
                ? (false, "Connected, but 'Product' table not found. Is this SimplePOSDB?")
                : (true,  $"✅ Connected → {conn.Database}  on  {conn.DataSource}");
        }
        catch (Exception ex)
        {
            return (false, $"❌ {ex.Message}");
        }
    }

    // ─── Main Import ──────────────────────────────────────────────────────────

    public async Task<ImportSummary> ImportAsync(
        List<ExcelProductRow>                               rows,
        IProgress<(int Current, int Total, string Message)> progress,
        CancellationToken cancellationToken = default)
    {
        var summary = new ImportSummary
        {
            TotalRows = rows.Count,
            StartTime = DateTime.Now
        };

        await using var conn = new SqlConnection(connectionString);
        await conn.OpenAsync(cancellationToken);

        // Pre-load lookup caches — avoids repeated round-trips for every row
        var brandCache    = await LoadLookupAsync(conn, "ProductBrands",     "BrandId",    "BrandName");
        var categoryCache = await LoadLookupAsync(conn, "ProductCategories", "CategoryId", "CategoryName");
        var groupCache    = await LoadLookupAsync(conn, "ProductGroups",     "GroupId",    "GroupName");

        int current = 0;

        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            current++;
            progress.Report((current, rows.Count,
                $"Processing {current}/{rows.Count}: {row.ProductCode} — {row.ProductName}"));

            // ── Skip invalid rows without touching the DB ─────────────────────
            if (!row.IsValid)
            {
                summary.Skipped++;
                summary.Results.Add(new ImportResult
                {
                    RowNumber   = row.RowNumber,
                    ProductCode = row.ProductCode,
                    ProductName = row.ProductName,
                    Status      = ImportStatus.Skipped,
                    Message     = row.ValidationMessage
                });
                continue;
            }

            // ── Each row is its own transaction ───────────────────────────────
            // One failure rolls back only that row — others are unaffected.
            await using var tx = conn.BeginTransaction();
            try
            {
                // Step 1 — ProductBrands
                int brandId = await GetOrInsertLookupAsync(conn, tx, brandCache,
                    "ProductBrands", "BrandId", "BrandName", row.ProductBrand);

                // Step 2 — ProductCategories
                int categoryId = await GetOrInsertLookupAsync(conn, tx, categoryCache,
                    "ProductCategories", "CategoryId", "CategoryName", row.Category);

                // Step 3 — ProductGroups  (default: "General")
                int groupId = await GetOrInsertLookupAsync(conn, tx, groupCache,
                    "ProductGroups", "GroupId", "GroupName", row.Group);

                // Step 4 — Product  (insert only if ProductCode not found)
                bool productExists = await ProductExistsAsync(conn, tx, row.ProductCode);
                if (!productExists)
                    await InsertProductAsync(conn, tx, row, brandId, categoryId, groupId);

                // Step 5 — Stock  (always — includes ProductSize)
                int stockId = await InsertStockAsync(conn, tx, row);

                // Step 6 — BillNumbers  (read → increment JournalVoucherNo → update)
                string voucherNumber = await GetNextJournalVoucherAsync(conn, tx);

                // Step 7 — StockTransaction  (uses the real voucher number from step 6)
                await InsertStockTransactionAsync(conn, tx, row, stockId, voucherNumber);

                // Step 8 — Voucher  (Opening Stock Balance, Debit 141, Credit 92)
                decimal amount    = row.TotalValue;  // Qty × CostPrice
                string  narration = $"Stock Opening Balance - Product Code: {row.ProductCode}";
                await InsertVoucherAsync(conn, tx, voucherNumber, amount, narration);

                // Step 9 — AccountsTransaction  (double-entry: 2 rows)
                await InsertAccountsTransactionAsync(conn, tx, voucherNumber, amount, narration);

                await tx.CommitAsync(cancellationToken);

                // ── Build result message ──────────────────────────────────────
                string sizeNote = string.IsNullOrWhiteSpace(row.ProductSize)
                    ? string.Empty
                    : $" | Size: {row.ProductSize}";

                string truncNote = row.ProductName.Length > 40
                    ? " | Name truncated"
                    : string.Empty;

                var status = productExists ? ImportStatus.ExistingProduct : ImportStatus.NewProduct;

                summary.Results.Add(new ImportResult
                {
                    RowNumber   = row.RowNumber,
                    ProductCode = row.ProductCode,
                    ProductName = row.ProductName,
                    Status      = status,
                    Message     = productExists
                        ? $"Stock added — {voucherNumber}{sizeNote}{truncNote}"
                        : $"New product + stock — {voucherNumber}{sizeNote}{truncNote}"
                });

                if (productExists) summary.ExistingProducts++;
                else               summary.NewProducts++;
            }
            catch (Exception ex)
            {
                await tx.RollbackAsync(cancellationToken);
                summary.Errors++;
                summary.Results.Add(new ImportResult
                {
                    RowNumber   = row.RowNumber,
                    ProductCode = row.ProductCode,
                    ProductName = row.ProductName,
                    Status      = ImportStatus.Error,
                    Message     = ex.Message
                });
            }
        }

        summary.EndTime = DateTime.Now;
        return summary;
    }

    // ─── Step 6: BillNumbers ─────────────────────────────────────────────────
    // Reads JournalVoucherNo → increments → updates → returns "JUR0000001"

    private static async Task<string> GetNextJournalVoucherAsync(
        SqlConnection conn, SqlTransaction tx)
    {
        await using var checkCmd = new SqlCommand(
            "SELECT COUNT(1) FROM [dbo].[BillNumbers]", conn, tx);
        bool hasRow = (int)(await checkCmd.ExecuteScalarAsync())! > 0;

        int nextNo;

        if (!hasRow)
        {
            // First ever use — create the counter row, all values 0 except JournalVoucherNo = 1
            const string insertSql = """
                INSERT INTO [dbo].[BillNumbers]
                    ([SalesBillNo],[SalesOrderBillNo],[SalesHoldNo],[SalesReturnNo],
                     [PurchaseBillNo],[PurchaseReturnNo],[ReceiptVoucherNo],[PaymentVoucherNo],
                     [JournalVoucherNo],[PurchaseVoucherNo],[PurchaseReturnVoucherNo],
                     [SalesVoucherNo],[SalesReturnVoucherNo])
                VALUES (0,0,0,0,0,0,0,0,1,0,0,0,0)
                """;
            await using var insertCmd = new SqlCommand(insertSql, conn, tx);
            await insertCmd.ExecuteNonQueryAsync();
            nextNo = 1;
        }
        else
        {
            await using var readCmd = new SqlCommand(
                "SELECT [JournalVoucherNo] FROM [dbo].[BillNumbers]", conn, tx);
            int current = Convert.ToInt32(await readCmd.ExecuteScalarAsync());
            nextNo = current + 1;

            await using var updateCmd = new SqlCommand(
                "UPDATE [dbo].[BillNumbers] SET [JournalVoucherNo] = @val", conn, tx);
            updateCmd.Parameters.AddWithValue("@val", nextNo);
            await updateCmd.ExecuteNonQueryAsync();
        }

        return $"JUR{nextNo:D7}";   // → JUR0000001 … JUR9999999
    }

    // ─── Step 8: Voucher ──────────────────────────────────────────────────────

    private static async Task InsertVoucherAsync(
        SqlConnection conn, SqlTransaction tx,
        string voucherNumber, decimal amount, string narration)
    {
        const string sql = """
            INSERT INTO [dbo].[Voucher]
                ([VoucherNumber], [VoucherDate], [VoucherType],
                 [DebitAccount],  [CreditAccount],
                 [TransactionAmount], [Narration])
            VALUES
                (@VoucherNumber, @VoucherDate, @VoucherType,
                 @DebitAccount,  @CreditAccount,
                 @Amount, @Narration)
            """;

        await using var cmd = new SqlCommand(sql, conn, tx);
        cmd.Parameters.AddWithValue("@VoucherNumber", voucherNumber);
        cmd.Parameters.AddWithValue("@VoucherDate",   DateOnly.FromDateTime(DateTime.Today));
        cmd.Parameters.AddWithValue("@VoucherType",   JournalVoucherType);
        cmd.Parameters.AddWithValue("@DebitAccount",  DebitAccountCode);    // 141
        cmd.Parameters.AddWithValue("@CreditAccount", CreditAccountCode);   // 92
        cmd.Parameters.AddWithValue("@Amount",        amount);
        cmd.Parameters.AddWithValue("@Narration",     narration);
        await cmd.ExecuteNonQueryAsync();
    }

    // ─── Step 9: AccountsTransaction — DEBIT + CREDIT rows ───────────────────

    private static async Task InsertAccountsTransactionAsync(
        SqlConnection conn, SqlTransaction tx,
        string voucherNumber, decimal amount, string narration)
    {
        const string sql = """
            INSERT INTO [dbo].[AccountsTransaction]
                ([TransactionDate], [AccountCode],
                 [DebitAmount], [CreditAmount],
                 [TransactionType], [VoucherNumber], [Narration], [UserID])
            VALUES
                (@Date, @AccountCode,
                 @Debit, @Credit,
                 @TxType, @VoucherNumber, @Narration, @UserID)
            """;

        // Row 1 — DEBIT 141 (Stock): inventory asset increases
        await using var debitCmd = new SqlCommand(sql, conn, tx);
        debitCmd.Parameters.AddWithValue("@Date",          DateOnly.FromDateTime(DateTime.Today));
        debitCmd.Parameters.AddWithValue("@AccountCode",   DebitAccountCode);
        debitCmd.Parameters.AddWithValue("@Debit",         amount);
        debitCmd.Parameters.AddWithValue("@Credit",        0m);
        debitCmd.Parameters.AddWithValue("@TxType",        AccountTxType);
        debitCmd.Parameters.AddWithValue("@VoucherNumber", voucherNumber);
        debitCmd.Parameters.AddWithValue("@Narration",     narration);
        debitCmd.Parameters.AddWithValue("@UserID",        ImportUserId);
        await debitCmd.ExecuteNonQueryAsync();

        // Row 2 — CREDIT 92 (Opening Balance Equity): equity increases
        await using var creditCmd = new SqlCommand(sql, conn, tx);
        creditCmd.Parameters.AddWithValue("@Date",          DateOnly.FromDateTime(DateTime.Today));
        creditCmd.Parameters.AddWithValue("@AccountCode",   CreditAccountCode);
        creditCmd.Parameters.AddWithValue("@Debit",         0m);
        creditCmd.Parameters.AddWithValue("@Credit",        amount);
        creditCmd.Parameters.AddWithValue("@TxType",        AccountTxType);
        creditCmd.Parameters.AddWithValue("@VoucherNumber", voucherNumber);
        creditCmd.Parameters.AddWithValue("@Narration",     narration);
        creditCmd.Parameters.AddWithValue("@UserID",        ImportUserId);
        await creditCmd.ExecuteNonQueryAsync();
    }

    // ─── Lookup helpers ───────────────────────────────────────────────────────

    private static async Task<Dictionary<string, int>> LoadLookupAsync(
        SqlConnection conn, string table, string idCol, string nameCol)
    {
        Dictionary<string, int> dict = new(StringComparer.OrdinalIgnoreCase);

        await using var cmd = new SqlCommand(
            $"SELECT [{idCol}], [{nameCol}] FROM [dbo].[{table}]", conn);
        await using var reader = await cmd.ExecuteReaderAsync();

        while (await reader.ReadAsync())
        {
            string name = reader[nameCol]?.ToString() ?? string.Empty;
            int    id   = Convert.ToInt32(reader[idCol]);
            if (!string.IsNullOrWhiteSpace(name))
                dict.TryAdd(name, id);
        }

        return dict;
    }

    private static async Task<int> GetOrInsertLookupAsync(
        SqlConnection conn, SqlTransaction tx,
        Dictionary<string, int> cache,
        string table, string idCol, string nameCol, string name)
    {
        if (string.IsNullOrWhiteSpace(name)) name = "General";
        if (cache.TryGetValue(name, out int existingId)) return existingId;

        string sql = $"""
            INSERT INTO [dbo].[{table}] ([{nameCol}])
            OUTPUT INSERTED.[{idCol}]
            VALUES (@name)
            """;

        await using var cmd = new SqlCommand(sql, conn, tx);
        cmd.Parameters.AddWithValue("@name", name);
        int newId = (int)(await cmd.ExecuteScalarAsync())!;
        cache[name] = newId;
        return newId;
    }

    private static async Task<bool> ProductExistsAsync(
        SqlConnection conn, SqlTransaction tx, string productCode)
    {
        await using var cmd = new SqlCommand(
            "SELECT COUNT(1) FROM [dbo].[Product] WHERE [ProductCode] = @code",
            conn, tx);
        cmd.Parameters.AddWithValue("@code", productCode);
        return (int)(await cmd.ExecuteScalarAsync())! > 0;
    }

    private static async Task InsertProductAsync(
        SqlConnection conn, SqlTransaction tx,
        ExcelProductRow row, int brandId, int categoryId, int groupId)
    {
        const string sql = """
            INSERT INTO [dbo].[Product]
                ([ProductCode], [CategoryId], [ProductName], [SalesPrice],
                 [SalesDiscount], [MinimumReorderQty], [ReorderLevel],
                 [GroupId], [BrandId], [IsActiveProduct])
            VALUES
                (@ProductCode, @CategoryId, @ProductName, @SalesPrice,
                 0, 0, 0, @GroupId, @BrandId, 1)
            """;

        await using var cmd = new SqlCommand(sql, conn, tx);
        cmd.Parameters.AddWithValue("@ProductCode", row.ProductCode);
        cmd.Parameters.AddWithValue("@CategoryId",  categoryId);
        cmd.Parameters.AddWithValue("@ProductName", row.ProductNameDb);   // already ≤40 chars
        cmd.Parameters.AddWithValue("@SalesPrice",  row.SellingPrice);
        cmd.Parameters.AddWithValue("@GroupId",     groupId);
        cmd.Parameters.AddWithValue("@BrandId",     brandId);
        await cmd.ExecuteNonQueryAsync();
    }

    private static async Task<int> InsertStockAsync(
        SqlConnection conn, SqlTransaction tx, ExcelProductRow row)
    {
        const string sql = """
            INSERT INTO [dbo].[Stock]
                ([ProductCode], [StockBalance], [ProductCost], [ProductPrice],
                 [ProductSize], [StockDate])
            OUTPUT INSERTED.[StockID]
            VALUES
                (@ProductCode, @StockBalance, @ProductCost, @ProductPrice,
                 @ProductSize, @StockDate)
            """;

        await using var cmd = new SqlCommand(sql, conn, tx);
        cmd.Parameters.AddWithValue("@ProductCode",  row.ProductCode);
        cmd.Parameters.AddWithValue("@StockBalance", row.Quantity);
        cmd.Parameters.AddWithValue("@ProductCost",  row.CostPrice);
        cmd.Parameters.AddWithValue("@ProductPrice", row.SellingPrice);
        cmd.Parameters.AddWithValue("@StockDate",    DateOnly.FromDateTime(DateTime.Today));

        // ProductSize — NULL in DB if not found in product name
        cmd.Parameters.AddWithValue("@ProductSize",
            string.IsNullOrWhiteSpace(row.ProductSize)
                ? (object)DBNull.Value
                : row.ProductSize);

        return (int)(await cmd.ExecuteScalarAsync())!;
    }

    private static async Task InsertStockTransactionAsync(
        SqlConnection conn, SqlTransaction tx,
        ExcelProductRow row, int stockId, string voucherNumber)
    {
        const string sql = """
            INSERT INTO [dbo].[StockTransaction]
                ([TransactionDate], [ProductCode], [StockID],
                 [StockReceived],   [StockCost],   [StockReceivedValue],
                 [StockIssued],     [StockIssuedValue],
                 [VoucherNumber],   [Narration],   [UserID])
            VALUES
                (@TransactionDate, @ProductCode, @StockID,
                 @StockReceived,   @StockCost,   @StockReceivedValue,
                 0, 0,
                 @VoucherNumber,   @Narration,   @UserID)
            """;

        await using var cmd = new SqlCommand(sql, conn, tx);
        cmd.Parameters.AddWithValue("@TransactionDate",    DateOnly.FromDateTime(DateTime.Today));
        cmd.Parameters.AddWithValue("@ProductCode",        row.ProductCode);
        cmd.Parameters.AddWithValue("@StockID",            stockId);
        cmd.Parameters.AddWithValue("@StockReceived",      row.Quantity);
        cmd.Parameters.AddWithValue("@StockCost",          row.CostPrice);
        cmd.Parameters.AddWithValue("@StockReceivedValue", row.TotalValue);
        cmd.Parameters.AddWithValue("@VoucherNumber",      voucherNumber);
        cmd.Parameters.AddWithValue("@Narration",
            $"Stock Opening Balance - Product Code: {row.ProductCode}");
        cmd.Parameters.AddWithValue("@UserID", ImportUserId);
        await cmd.ExecuteNonQueryAsync();
    }
}
