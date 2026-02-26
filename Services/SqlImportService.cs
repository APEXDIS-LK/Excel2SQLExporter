using Microsoft.Data.SqlClient;
using Excel2SQLExporter.Models;

namespace Excel2SQLExporter.Services;

/// <summary>
/// Handles all SQL Server import operations for SimplePOSDB.
///
/// C# 14: Primary constructor.
///
/// VoucherMode.PerProduct  — each row gets its own JUR number.
///                           Voucher + AccountsTransaction written inside the per-row transaction.
///
/// VoucherMode.BatchSingle — one JUR number for the whole import.
///                           BillNumbers incremented ONCE before the loop.
///                           Per-row transactions cover Product / Stock / StockTransaction only.
///                           ONE Voucher + 2 AccountsTransaction rows written in a final transaction.
/// </summary>
public class SqlImportService(string connectionString)
{
    private const string DebitAccountCode   = "141";                 // Stock (Asset)
    private const string CreditAccountCode  = "92";                  // Opening Balance Equity
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

            int count = (int)(await cmd.ExecuteScalarAsync())!;
            return count == 0
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
        VoucherMode                                         voucherMode,
        bool                                                insertBarcodes,
        IProgress<(int Current, int Total, string Message)> progress,
        CancellationToken cancellationToken = default)
    {
        var summary = new ImportSummary
        {
            TotalRows       = rows.Count,
            StartTime       = DateTime.Now,
            VoucherMode     = voucherMode,
            BarcodesEnabled = insertBarcodes
        };

        await using var conn = new SqlConnection(connectionString);
        await conn.OpenAsync(cancellationToken);

        var brandCache    = await LoadLookupAsync(conn, "ProductBrands",     "BrandId",    "BrandName");
        var categoryCache = await LoadLookupAsync(conn, "ProductCategories", "CategoryId", "CategoryName");
        var groupCache    = await LoadLookupAsync(conn, "ProductGroups",     "GroupId",    "GroupName");

        // Load the price cipher key from RegisteredUser.CostCode once — used by all barcode rows.
        // Falls back to null if the table is empty or the column is blank; barcodes will
        // then write NULL to PrintBarCodes.CostCode rather than failing the import.
        string? priceKey = insertBarcodes
            ? await LoadPriceKeyAsync(conn)
            : null;

        // ── Batch mode pre-step: reserve ONE voucher number ───────────────────
        string? batchVoucherNo   = null;
        decimal batchTotalValue  = 0m;
        int     batchValidCount  = 0;

        if (voucherMode == VoucherMode.BatchSingle)
        {
            progress.Report((0, rows.Count, "📋  Batch mode — reserving a single Journal Voucher number..."));
            await using var preTx = conn.BeginTransaction();
            try
            {
                batchVoucherNo = await GetNextJournalVoucherAsync(conn, preTx);
                await preTx.CommitAsync(cancellationToken);
                progress.Report((0, rows.Count, $"📋  Batch voucher reserved: {batchVoucherNo}"));
            }
            catch
            {
                await preTx.RollbackAsync(cancellationToken);
                throw;
            }

            // Pre-calculate total batch value and count from valid rows
            batchTotalValue = rows.Where(r => r.IsValid).Sum(r => r.TotalValue);
            batchValidCount = rows.Count(r => r.IsValid);

            summary.BatchVoucherNo  = batchVoucherNo;
            summary.BatchTotalValue = batchTotalValue;
        }

        // ── Per-row loop ──────────────────────────────────────────────────────
        int current = 0;

        foreach (var row in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            current++;
            progress.Report((current, rows.Count,
                $"Processing {current}/{rows.Count}: {row.ProductCode} — {row.ProductName}"));

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

            await using var tx = conn.BeginTransaction();
            try
            {
                int brandId = await GetOrInsertLookupAsync(conn, tx, brandCache,
                    "ProductBrands", "BrandId", "BrandName", row.ProductBrand);

                int categoryId = await GetOrInsertLookupAsync(conn, tx, categoryCache,
                    "ProductCategories", "CategoryId", "CategoryName", row.Category);

                int groupId = await GetOrInsertLookupAsync(conn, tx, groupCache,
                    "ProductGroups", "GroupId", "GroupName", row.Group);

                bool productExists = await ProductExistsAsync(conn, tx, row.ProductCode);
                if (!productExists)
                    await InsertProductAsync(conn, tx, row, brandId, categoryId, groupId);

                int stockId = await InsertStockAsync(conn, tx, row);

                // ── Voucher number depends on mode ────────────────────────────
                string voucherNumber;

                if (voucherMode == VoucherMode.PerProduct)
                {
                    voucherNumber = await GetNextJournalVoucherAsync(conn, tx);
                    await InsertStockTransactionAsync(conn, tx, row, stockId, voucherNumber);
                    string narration = $"Stock Opening Balance - Product Code: {row.ProductCode}";
                    await InsertVoucherAsync(conn, tx, voucherNumber, row.TotalValue, narration);
                    await InsertAccountsTransactionAsync(conn, tx, voucherNumber, row.TotalValue, narration);
                }
                else
                {
                    voucherNumber = batchVoucherNo!;
                    await InsertStockTransactionAsync(conn, tx, row, stockId, voucherNumber);
                }

                // ── Barcodes: insert (int)Quantity rows into PrintBarCodes ────
                int barcodesWritten = 0;
                if (insertBarcodes)
                {
                    barcodesWritten = await InsertBarcodesAsync(conn, tx, row, priceKey);
                    summary.TotalBarcodesInserted += barcodesWritten;
                }

                await tx.CommitAsync(cancellationToken);

                string sizeNote    = string.IsNullOrWhiteSpace(row.ProductSize) ? string.Empty : $" | Size: {row.ProductSize}";
                string truncNote   = row.ProductName.Length > 40 ? " | Name truncated" : string.Empty;
                string barcodeNote = insertBarcodes ? $" | Barcodes: {barcodesWritten}" : string.Empty;
                var    status      = productExists ? ImportStatus.ExistingProduct : ImportStatus.NewProduct;

                summary.Results.Add(new ImportResult
                {
                    RowNumber   = row.RowNumber,
                    ProductCode = row.ProductCode,
                    ProductName = row.ProductName,
                    Status      = status,
                    Message     = productExists
                        ? $"Stock added — {voucherNumber}{sizeNote}{truncNote}{barcodeNote}"
                        : $"New product + stock — {voucherNumber}{sizeNote}{truncNote}{barcodeNote}"
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

        // ── Batch mode post-step: write ONE Voucher + 2 AccountsTransaction ───
        if (voucherMode == VoucherMode.BatchSingle && batchVoucherNo is not null)
        {
            int  successCount = summary.NewProducts + summary.ExistingProducts;
            string batchNarration =
                $"Stock Opening Balance - Batch Import {DateTime.Today:dd/MM/yyyy} ({successCount} products)";

            progress.Report((rows.Count, rows.Count,
                $"📒  Writing batch Voucher {batchVoucherNo}  (total: {batchTotalValue:N2}, {successCount} products)..."));

            await using var postTx = conn.BeginTransaction();
            try
            {
                await InsertVoucherAsync(postTx.Connection!, postTx,
                    batchVoucherNo, batchTotalValue, batchNarration);
                await InsertAccountsTransactionAsync(postTx.Connection!, postTx,
                    batchVoucherNo, batchTotalValue, batchNarration);
                await postTx.CommitAsync(cancellationToken);

                progress.Report((rows.Count, rows.Count,
                    $"✅  Batch Voucher {batchVoucherNo} committed — Debit 141 / Credit 92 for {batchTotalValue:N2}"));
            }
            catch (Exception ex)
            {
                await postTx.RollbackAsync(cancellationToken);
                progress.Report((rows.Count, rows.Count,
                    $"❌  Batch Voucher write failed: {ex.Message}"));
            }
        }

        summary.EndTime = DateTime.Now;
        return summary;
    }

    // ─── BillNumbers: read → increment JournalVoucherNo → update → return "JUR0000001" ──

    private static async Task<string> GetNextJournalVoucherAsync(
        SqlConnection conn, SqlTransaction tx)
    {
        await using var checkCmd = new SqlCommand(
            "SELECT COUNT(1) FROM [dbo].[BillNumbers]", conn, tx);
        bool hasRow = (int)(await checkCmd.ExecuteScalarAsync())! > 0;

        int nextNo;

        if (!hasRow)
        {
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

        return $"JUR{nextNo:D7}";
    }

    // ─── Voucher ──────────────────────────────────────────────────────────────

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
        cmd.Parameters.AddWithValue("@DebitAccount",  DebitAccountCode);
        cmd.Parameters.AddWithValue("@CreditAccount", CreditAccountCode);
        cmd.Parameters.AddWithValue("@Amount",        amount);
        cmd.Parameters.AddWithValue("@Narration",     narration);
        await cmd.ExecuteNonQueryAsync();
    }

    // ─── AccountsTransaction — DEBIT + CREDIT ────────────────────────────────

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

        // DEBIT 141 — Stock: inventory asset increases
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

        // CREDIT 92 — Opening Balance Equity: equity increases
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

    /// <summary>
    /// Reads CostCode from RegisteredUser (the price cipher key, e.g. "HOLYQURANF").
    /// Returns null if the table is empty or CostCode is blank — callers handle null gracefully.
    /// </summary>
    private static async Task<string?> LoadPriceKeyAsync(SqlConnection conn)
    {
        await using var cmd = new SqlCommand(
            "SELECT TOP 1 [CostCode] FROM [dbo].[RegisteredUser] WHERE [CostCode] IS NOT NULL AND LEN([CostCode]) >= 10",
            conn);
        var result = await cmd.ExecuteScalarAsync();
        return result is string s && s.Length >= 10 ? s : null;
    }

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
                ([ProductCode], [BarCode], [CategoryId], [ProductName], [SalesPrice],
                 [SalesDiscount], [MinimumReorderQty], [ReorderLevel],
                 [GroupId], [BrandId], [IsActiveProduct])
            VALUES
                (@ProductCode, @BarCode, @CategoryId, @ProductName, @SalesPrice,
                 0, 0, 0, @GroupId, @BrandId, 1)
            """;

        // BarCode column is nvarchar(13) — use first 13 chars of ProductCode
        string barCode = row.ProductCode.Length > 13
            ? row.ProductCode[..13]
            : row.ProductCode;

        await using var cmd = new SqlCommand(sql, conn, tx);
        cmd.Parameters.AddWithValue("@ProductCode", row.ProductCode);
        cmd.Parameters.AddWithValue("@BarCode",     barCode);
        cmd.Parameters.AddWithValue("@CategoryId",  categoryId);
        cmd.Parameters.AddWithValue("@ProductName", row.ProductNameDb);
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

    // ─── PrintBarCodes: insert one row per unit of stock ─────────────────────
    // If Qty = 4, this inserts 4 rows — one per physical label to be printed.
    //
    // Column constraints (from schema):
    //   BarCodeString  nvarchar(13)  — ProductCode truncated to 13
    //   Description    nvarchar(20)  — ProductNameDb truncated to 20
    //   Price          nvarchar(12)  — SellingPrice as "F2" string
    //   CostCode       nvarchar(10)  — encoded via PriceCodeConverter (RegisteredUser.CostCode)
    //                                  NULL if priceKey not found or shorter than 10 chars
    //   ProductSize    nvarchar(10)  — extracted size, NULL if empty
    //   ProductColor   nvarchar(10)  — NULL
    //   Remarks        nvarchar(10)  — NULL
    //   Remarks1       nvarchar(10)  — NULL

    private static async Task<int> InsertBarcodesAsync(
        SqlConnection conn, SqlTransaction tx,
        ExcelProductRow row, string? priceKey)
    {
        const string sql = """
            INSERT INTO [dbo].[PrintBarCodes]
                ([BarCodeString], [Description], [Price],
                 [CostCode], [ProductSize], [ProductColor],
                 [Remarks], [Remarks1])
            VALUES
                (@BarCodeString, @Description, @Price,
                 @CostCode, @ProductSize, NULL,
                 NULL, NULL)
            """;

        // Enforce nvarchar length limits from schema
        string barCodeStr  = row.ProductCode.Length  > 13 ? row.ProductCode[..13]   : row.ProductCode;
        string description = row.ProductNameDb.Length > 20 ? row.ProductNameDb[..20] : row.ProductNameDb;
        string price       = row.SellingPrice.ToString("F2");
        if (price.Length > 12) price = price[..12];

        string? productSize = string.IsNullOrWhiteSpace(row.ProductSize) ? null
            : row.ProductSize.Length > 10 ? row.ProductSize[..10]
            : row.ProductSize;

        // Compute CostCode: encode the price using the cipher key from RegisteredUser.CostCode.
        // priceKey is null when the table is empty or the column is blank → CostCode stays NULL.
        string? costCode = null;
        if (priceKey is not null)
        {
            string raw = PriceCodeConverter.ConvertToCostCode(price, priceKey);
            costCode   = raw.Length > 10 ? raw[..10] : raw;   // nvarchar(10) limit
        }

        // One row per unit of stock — Qty 4 → 4 label rows
        int count = (int)Math.Max(1, Math.Floor(row.Quantity));

        for (int i = 0; i < count; i++)
        {
            await using var cmd = new SqlCommand(sql, conn, tx);
            cmd.Parameters.AddWithValue("@BarCodeString", barCodeStr);
            cmd.Parameters.AddWithValue("@Description",   description);
            cmd.Parameters.AddWithValue("@Price",         price);
            cmd.Parameters.AddWithValue("@CostCode",      costCode    is null ? (object)DBNull.Value : costCode);
            cmd.Parameters.AddWithValue("@ProductSize",   productSize is null ? (object)DBNull.Value : productSize);
            await cmd.ExecuteNonQueryAsync();
        }

        return count;
    }
}
