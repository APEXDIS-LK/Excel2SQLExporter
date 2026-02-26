using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using Excel2SQLExporter.Models;
using Excel2SQLExporter.Services;
using Microsoft.Win32;

namespace Excel2SQLExporter;

/// <summary>
/// Converts bool (IsValid) to ✅ / ⚠️ in the DataGrid Valid column.
/// </summary>
public class BoolToTickConverter : IValueConverter
{
    public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        => value is true ? "✅" : "⚠️";

    public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        => throw new NotImplementedException();
}

public partial class MainWindow : Window
{
    // C# 14: field keyword — lazy-initialised reader, no separate backing field
    private ExcelReaderService ExcelReader
    {
        get => field ??= new ExcelReaderService();
    }

    private List<ExcelProductRow> _previewRows = [];
    private CancellationTokenSource? _cts;
    private ImportSummary? _lastSummary;

    // Currently selected voucher mode — driven by radio buttons
    private VoucherMode SelectedVoucherMode
        => RadioBatchSingle.IsChecked == true
            ? VoucherMode.BatchSingle
            : VoucherMode.PerProduct;

    // Whether to insert PrintBarCodes rows — driven by the checkbox
    private bool InsertBarcodes => ChkInsertBarcodes.IsChecked == true;

    public MainWindow()
    {
        InitializeComponent();

        // Set window icon from embedded resource — more reliable than XAML Icon= attribute
        // which can fail with BAML pack URI resolution depending on build configuration.
        try
        {
            var uri    = new Uri("pack://application:,,,/Excel2SQLExporter;component/app.ico");
            var stream = Application.GetResourceStream(uri);
            if (stream != null)
                Icon = System.Windows.Media.Imaging.BitmapFrame.Create(
                    stream.Stream,
                    System.Windows.Media.Imaging.BitmapCreateOptions.None,
                    System.Windows.Media.Imaging.BitmapCacheOption.OnLoad);
        }
        catch
        {
            // Non-fatal — app runs fine without a custom window icon
        }

        Log("👋  Ready. Set connection string → select Excel file → choose voucher mode → Preview → Import.");
        Log("ℹ️   Per-product mode: one JUR per row.  Batch mode: one JUR shared by all rows.");
    }

    // ─── VOUCHER MODE RADIO ───────────────────────────────────────────────────

    private void VoucherModeChanged(object sender, RoutedEventArgs e)
    {
        // Guard: may fire before InitializeComponent completes
        if (TxtVoucherModeBadge is null || BadgeVoucherMode is null) return;

        if (SelectedVoucherMode == VoucherMode.BatchSingle)
        {
            TxtVoucherModeBadge.Text           = "● Batch Single";
            BadgeVoucherMode.Background        = new SolidColorBrush(Color.FromRgb(240, 253, 244));
            TxtVoucherModeBadge.Foreground     = new SolidColorBrush(Color.FromRgb(21,  128, 61));
        }
        else
        {
            TxtVoucherModeBadge.Text           = "● Per Product";
            BadgeVoucherMode.Background        = new SolidColorBrush(Color.FromRgb(239, 246, 255));
            TxtVoucherModeBadge.Foreground     = new SolidColorBrush(Color.FromRgb(37,  99,  235));
        }

        // Hide batch voucher badge from a previous run when mode changes
        BadgeBatchVoucher.Visibility = Visibility.Collapsed;

        Log($"📒  Voucher mode: {SelectedVoucherMode switch {
            VoucherMode.PerProduct  => "One voucher per product row",
            VoucherMode.BatchSingle => "One voucher for entire batch",
            _                       => "Unknown"
        }}");
    }

    // ─── BARCODE CHECKBOX ─────────────────────────────────────────────────────

    private void ChkInsertBarcodes_Changed(object sender, RoutedEventArgs e)
    {
        if (BadgeBarcodeOption is null) return;

        if (InsertBarcodes)
        {
            BadgeBarcodeOption.Visibility = Visibility.Visible;
            Log("🏷️  Barcode insert ON — PrintBarCodes will receive Qty rows per product.");
        }
        else
        {
            BadgeBarcodeOption.Visibility = Visibility.Collapsed;
            Log("🏷️  Barcode insert OFF — PrintBarCodes will not be written.");
        }
    }

    // ─── TEST CONNECTION ──────────────────────────────────────────────────────

    private async void BtnTestConnection_Click(object sender, RoutedEventArgs e)
    {
        TxtConnectionStatus.Text       = "Testing...";
        TxtConnectionStatus.Foreground = Brushes.Gray;

        var (success, message) = await new SqlImportService(
            TxtConnectionString.Text.Trim()).TestConnectionAsync();

        TxtConnectionStatus.Text       = message;
        TxtConnectionStatus.Foreground = success
            ? new SolidColorBrush(Color.FromRgb(22,  163, 74))
            : new SolidColorBrush(Color.FromRgb(220, 38,  38));

        Log(message);
    }

    // ─── BROWSE ───────────────────────────────────────────────────────────────

    private void BtnBrowse_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Title  = "Select Excel File",
            Filter = "Excel Files (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|All Files (*.*)|*.*"
        };

        if (dlg.ShowDialog() != true) return;

        TxtFilePath.Text = dlg.FileName;
        Log($"📂  File selected: {Path.GetFileName(dlg.FileName)}");

        _previewRows              = [];
        PreviewGrid.ItemsSource   = null;
        TxtRowCount.Text          = "0 rows";
        BtnImport.IsEnabled       = false;
        BtnRecall.IsEnabled       = false;
        BtnImportSelected.IsEnabled = false;
        BadgeSelected.Visibility  = Visibility.Collapsed;
        BadgeWarnings.Visibility  = Visibility.Collapsed;
    }

    // ─── PREVIEW ──────────────────────────────────────────────────────────────

    private void BtnPreview_Click(object sender, RoutedEventArgs e)
    {
        string filePath = TxtFilePath.Text.Trim();

        if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
        {
            MessageBox.Show("Please select a valid Excel file first.", "No File Selected",
                MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        try
        {
            Log("🔍  Reading Excel file...");
            var (rows, warnings) = ExcelReader.ReadExcel(filePath);

            _previewRows            = rows;
            PreviewGrid.ItemsSource = new ObservableCollection<ExcelProductRow>(rows);

            int validCount   = rows.Count(r => r.IsValid);
            int invalidCount = rows.Count(r => !r.IsValid);
            int withSize     = rows.Count(r => !string.IsNullOrWhiteSpace(r.ProductSize));
            int truncated    = rows.Count(r => r.ProductName.Length > 40);
            decimal batchTotal = rows.Where(r => r.IsValid).Sum(r => r.TotalValue);

            TxtRowCount.Text = $"{rows.Count} rows";

            if (warnings.Count > 0 || invalidCount > 0)
            {
                BadgeWarnings.Visibility = Visibility.Visible;
                TxtWarningCount.Text     = $"⚠️ {warnings.Count + invalidCount} warnings";
                warnings.ForEach(Log);
            }
            else
            {
                BadgeWarnings.Visibility = Visibility.Collapsed;
            }

            Log($"✅  Preview loaded — {rows.Count} rows:");
            Log($"     ✅ Valid          : {validCount}");
            Log($"     ⚠️  Invalid/skipped: {invalidCount}");
            Log($"     📐 With size      : {withSize}");
            Log($"     ✂️  Name >40 chars  : {truncated}  (will be truncated in DB)");
            Log($"     💰 Batch total    : {batchTotal:N2}");

            BtnImport.IsEnabled = validCount > 0;
            BtnRecall.IsEnabled = validCount > 0;
        }
        catch (Exception ex)
        {
            Log($"❌  Failed to read file: {ex.Message}");
            MessageBox.Show($"Could not read the Excel file:\n\n{ex.Message}",
                "Read Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }


    // ─── SELECTION CHANGED ────────────────────────────────────────────────────

    private void PreviewGrid_SelectionChanged(object sender,
        System.Windows.Controls.SelectionChangedEventArgs e)
    {
        int total = PreviewGrid.SelectedItems.Count;
        int valid = PreviewGrid.SelectedItems
            .Cast<ExcelProductRow>()
            .Count(r => r.IsValid);

        if (total > 0)
        {
            BadgeSelected.Visibility = Visibility.Visible;
            TxtSelectedCount.Text    = valid < total
                ? $"{total} selected  ({valid} valid)"
                : $"{total} selected";
        }
        else
        {
            BadgeSelected.Visibility = Visibility.Collapsed;
        }

        // Enable only when at least one valid row is highlighted and not mid-import
        BtnImportSelected.IsEnabled = valid > 0 && !BtnCancel.IsEnabled;
    }

    // ─── IMPORT ALL ───────────────────────────────────────────────────────────

    private async void BtnImport_Click(object sender, RoutedEventArgs e)
    {
        if (_previewRows.Count == 0)
        {
            MessageBox.Show("No data to import. Please preview the file first.",
                "No Data", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        if (!ConfirmImport(_previewRows, "all valid rows")) return;
        await ExecuteImportAsync(_previewRows, "All rows");
    }

    // ─── IMPORT SELECTED ─────────────────────────────────────────────────────

    private async void BtnImportSelected_Click(object sender, RoutedEventArgs e)
    {
        if (PreviewGrid.SelectedItems.Count == 0)
        {
            MessageBox.Show(
                "No rows are selected in the preview grid.\n\n" +
                "Click a row to select it, Ctrl+Click for individual rows, Shift+Click for a range.",
                "Nothing Selected", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var rowsToImport = PreviewGrid.SelectedItems
            .Cast<ExcelProductRow>()
            .ToList();

        int validCount = rowsToImport.Count(r => r.IsValid);
        if (validCount == 0)
        {
            MessageBox.Show(
                "None of the selected rows are valid (check the ✓ column).\n\n" +
                "Only valid rows can be imported.",
                "No Valid Rows", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        string scope = rowsToImport.Count == validCount
            ? $"{validCount} selected rows"
            : $"{validCount} valid of {rowsToImport.Count} selected rows";

        if (!ConfirmImport(rowsToImport, scope)) return;
        await ExecuteImportAsync(rowsToImport, $"Selected ({validCount} rows)");
    }

    // ─── SHARED: confirm dialog ───────────────────────────────────────────────

    private bool ConfirmImport(List<ExcelProductRow> rows, string scope)
    {
        string connStr = TxtConnectionString.Text.Trim();
        if (string.IsNullOrWhiteSpace(connStr))
        {
            MessageBox.Show("Please enter a SQL Server connection string.",
                "No Connection", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        var     mode        = SelectedVoucherMode;
        bool    barcodes    = InsertBarcodes;
        int     validCount  = rows.Count(r => r.IsValid);
        decimal batchTotal  = rows.Where(r => r.IsValid).Sum(r => r.TotalValue);
        int     totalLabels = barcodes
            ? rows.Where(r => r.IsValid).Sum(r => (int)Math.Floor(r.Quantity))
            : 0;

        string modeDesc = mode == VoucherMode.PerProduct
            ? $"  • {validCount} separate JUR numbers (one per row)\n" +
              $"  • {validCount} Voucher records\n" +
              $"  • {validCount * 2} AccountsTransaction rows"
            : $"  • 1 shared JUR number for all rows\n" +
              $"  • 1 Voucher record (total: {batchTotal:N2})\n" +
              $"  • 2 AccountsTransaction rows";

        string barcodeDesc = barcodes
            ? $"\n\n🏷️  Barcodes: {totalLabels} label rows → PrintBarCodes"
            : string.Empty;

        var confirm = MessageBox.Show(
            $"Ready to import {validCount} valid rows  ({scope}).\n\n" +
            $"Voucher Mode: {(mode == VoucherMode.PerProduct ? "ONE VOUCHER PER PRODUCT" : "ONE VOUCHER FOR ENTIRE BATCH")}\n\n" +
            $"{modeDesc}\n\n" +
            "Each row also writes: Product (if new) · Stock · StockTransaction" +
            barcodeDesc + "\n\n" +
            "Per-row failures are isolated — they won't affect other rows.\n\nContinue?",
            "Confirm Import",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        return confirm == MessageBoxResult.Yes;
    }

    // ─── SHARED: run import, update all UI ───────────────────────────────────

    private async Task ExecuteImportAsync(List<ExcelProductRow> rowsToImport, string label)
    {
        string  connStr    = TxtConnectionString.Text.Trim();
        var     mode       = SelectedVoucherMode;
        bool    barcodes   = InsertBarcodes;
        int     validCount = rowsToImport.Count(r => r.IsValid);
        decimal batchTotal = rowsToImport.Where(r => r.IsValid).Sum(r => r.TotalValue);
        int     totalLabels = barcodes
            ? rowsToImport.Where(r => r.IsValid).Sum(r => (int)Math.Floor(r.Quantity))
            : 0;

        SetImportingState(true);
        ResetBadges();
        BadgeBatchVoucher.Visibility = Visibility.Collapsed;
        ImportLog.Items.Clear();
        ImportProgress.Value  = 0;
        TxtProgressLabel.Text = "0%";

        Log($"🚀  Import started — {DateTime.Now:HH:mm:ss}");
        Log($"    Scope   : {label}  ({validCount} valid rows)");
        Log($"    Mode    : {(mode == VoucherMode.PerProduct ? "Per Product" : "Batch Single")}");
        Log($"    Total   : {batchTotal:N2}");
        if (barcodes) Log($"    Labels  : {totalLabels} barcode rows → PrintBarCodes");

        _cts = new CancellationTokenSource();

        var progress = new Progress<(int Current, int Total, string Message)>(p =>
        {
            if (p.Total > 0)
            {
                double pct            = (double)p.Current / p.Total * 100;
                ImportProgress.Value  = Math.Min(100, pct);
                TxtProgressLabel.Text = $"{pct:F0}%";
            }
            Log(p.Message);
            ScrollLog();
        });

        try
        {
            _lastSummary = await new SqlImportService(connStr)
                .ImportAsync(rowsToImport, mode, barcodes, progress, _cts.Token);

            TxtNewCount.Text     = $"✅ New: {_lastSummary.NewProducts}";
            TxtStockCount.Text   = $"🔄 Stock Added: {_lastSummary.ExistingProducts}";
            TxtSkippedCount.Text = $"⚠️ Skipped: {_lastSummary.Skipped}";
            TxtErrorCount.Text   = $"❌ Errors: {_lastSummary.Errors}";
            ImportProgress.Value  = 100;
            TxtProgressLabel.Text = "100%";

            if (_lastSummary.BarcodesEnabled)
            {
                TxtBarcodeCount.Text         = $"🏷️ Labels: {_lastSummary.TotalBarcodesInserted}";
                BadgeBarcodeCount.Visibility = Visibility.Visible;
            }
            else
            {
                BadgeBarcodeCount.Visibility = Visibility.Collapsed;
            }

            if (mode == VoucherMode.BatchSingle && !string.IsNullOrEmpty(_lastSummary.BatchVoucherNo))
            {
                TxtBatchVoucherNo.Text       = $"📒 Batch: {_lastSummary.BatchVoucherNo}  ({_lastSummary.BatchTotalValue:N2})";
                BadgeBatchVoucher.Visibility = Visibility.Visible;
            }

            Log("─────────────────────────────────────────────────");
            Log($"✅  Complete — {_lastSummary.EndTime:HH:mm:ss}");
            Log($"    New products  : {_lastSummary.NewProducts}");
            Log($"    Stock added   : {_lastSummary.ExistingProducts}");
            Log($"    Skipped       : {_lastSummary.Skipped}");
            Log($"    Errors        : {_lastSummary.Errors}");
            if (_lastSummary.BarcodesEnabled)
                Log($"    Label rows    : {_lastSummary.TotalBarcodesInserted}  (PrintBarCodes)");
            if (mode == VoucherMode.BatchSingle)
                Log($"    Batch voucher : {_lastSummary.BatchVoucherNo}  (total: {_lastSummary.BatchTotalValue:N2})");
            Log($"    Duration      : {_lastSummary.Duration.TotalSeconds:F1}s");
            Log("─────────────────────────────────────────────────");

            if (_lastSummary.Errors > 0)
            {
                Log("❌  Error details:");
                foreach (var r in _lastSummary.Results.Where(r => r.Status == ImportStatus.Error))
                    Log($"     Row {r.RowNumber} [{r.ProductCode}]: {r.Message}");
            }

            string barcodeLine = _lastSummary.BarcodesEnabled
                ? $"\n🏷️  Label rows     : {_lastSummary.TotalBarcodesInserted}  (PrintBarCodes)"
                : string.Empty;

            string batchLine = mode == VoucherMode.BatchSingle
                ? $"\n📒  Batch Voucher  : {_lastSummary.BatchVoucherNo}  ({_lastSummary.BatchTotalValue:N2})"
                : string.Empty;

            MessageBox.Show(
                $"Import complete!  [{label}]\n\n" +
                $"✅  New products  : {_lastSummary.NewProducts}\n" +
                $"🔄  Stock added   : {_lastSummary.ExistingProducts}\n" +
                $"⚠️   Skipped       : {_lastSummary.Skipped}\n" +
                $"❌  Errors        : {_lastSummary.Errors}" +
                barcodeLine + batchLine +
                $"\n\nDuration: {_lastSummary.Duration.TotalSeconds:F1}s\n\n" +
                "Use 'Export Log as CSV' to save a detailed report.",
                "Import Complete",
                MessageBoxButton.OK,
                _lastSummary.Errors > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
        }
        catch (OperationCanceledException)
        {
            Log("⛔  Import cancelled.");
            ImportProgress.Value  = 0;
            TxtProgressLabel.Text = "0%";
        }
        catch (Exception ex)
        {
            Log($"❌  Import failed: {ex.Message}");
            MessageBox.Show($"Import failed:\n\n{ex.Message}",
                "Import Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetImportingState(false);
            _cts?.Dispose();
            _cts = null;
        }
    }



    // ─── CANCEL ───────────────────────────────────────────────────────────────

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        _cts?.Cancel();
        Log("⛔  Cancellation requested — finishing current row then stopping...");
    }

    // ─── RECALL ───────────────────────────────────────────────────────────────

    private async void BtnRecall_Click(object sender, RoutedEventArgs e)
    {
        if (_previewRows.Count == 0)
        {
            MessageBox.Show("Please preview an Excel file first.\nRecall needs a list of ProductCodes to match against.",
                "No Preview Loaded", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        var validRows   = _previewRows.Where(r => r.IsValid).ToList();
        var productCodes = validRows.Select(r => r.ProductCode).Distinct().ToList();

        var confirm = MessageBox.Show(
            $"⚠️  RECALL — Delete matching DB records\n\n" +
            $"This will DELETE all records in the following tables\n" +
            $"where ProductCode matches one of the {productCodes.Count} codes in the current Excel file:\n\n" +
            $"  • Product\n" +
            $"  • Stock\n" +
            $"  • StockTransaction\n" +
            $"  • Voucher  (via StockTransaction VoucherNumber)\n" +
            $"  • AccountsTransaction  (via same VoucherNumbers)\n" +
            $"  • PrintBarCodes\n\n" +
            $"Lookup tables (Brands, Categories, Groups) are NOT touched.\n\n" +
            $"This cannot be undone. Continue?",
            "Confirm Recall",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (confirm != MessageBoxResult.Yes) return;

        SetImportingState(true);
        TxtDeleteStatus.Text      = "Running Recall...";
        TxtDeleteStatus.Foreground = System.Windows.Media.Brushes.DarkOrange;
        Log("─────────────────────────────────────────────────");
        Log($"🗑️  RECALL started — {DateTime.Now:HH:mm:ss}");
        Log($"    Matching {productCodes.Count} ProductCodes from current Excel file");

        var progress = new Progress<(int Current, int Total, string Message)>(p =>
        {
            double pct = p.Total > 0 ? (double)p.Current / p.Total * 100 : 0;
            ImportProgress.Value  = pct;
            TxtProgressLabel.Text = $"{pct:F0}%";
            Log(p.Message);
            ScrollLog();
        });

        try
        {
            var svc     = new SqlDeleteService(TxtConnectionString.Text.Trim());
            var summary = await svc.RecallByProductCodesAsync(productCodes, progress);

            ImportProgress.Value  = 100;
            TxtProgressLabel.Text = "100%";

            string statusText =
                $"Recall: Products={summary.DeletedProducts}  " +
                $"Stock={summary.DeletedStockRecords}  " +
                $"StockTx={summary.DeletedStockTransactions}  " +
                $"Vouchers={summary.DeletedVouchers}  " +
                $"AcctTx={summary.DeletedAccountsTransactions}  " +
                $"Barcodes={summary.DeletedBarcodes}  " +
                $"({summary.Duration.TotalSeconds:F1}s)";

            TxtDeleteStatus.Text       = statusText;
            TxtDeleteStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(21, 128, 61));

            Log("─────────────────────────────────────────────────");
            Log($"✅  Recall complete — {summary.Duration.TotalSeconds:F1}s");
            Log($"    Products deleted          : {summary.DeletedProducts}");
            Log($"    Stock rows deleted        : {summary.DeletedStockRecords}");
            Log($"    StockTransaction rows     : {summary.DeletedStockTransactions}");
            Log($"    Voucher rows deleted      : {summary.DeletedVouchers}");
            Log($"    AccountsTransaction rows  : {summary.DeletedAccountsTransactions}");
            Log($"    PrintBarCodes rows        : {summary.DeletedBarcodes}");
            Log("─────────────────────────────────────────────────");

            MessageBox.Show(
                $"Recall complete!\n\n" +
                $"Products deleted          : {summary.DeletedProducts}\n" +
                $"Stock rows deleted        : {summary.DeletedStockRecords}\n" +
                $"StockTransaction rows     : {summary.DeletedStockTransactions}\n" +
                $"Voucher rows              : {summary.DeletedVouchers}\n" +
                $"AccountsTransaction rows  : {summary.DeletedAccountsTransactions}\n" +
                $"PrintBarCodes rows        : {summary.DeletedBarcodes}\n\n" +
                $"Duration: {summary.Duration.TotalSeconds:F1}s",
                "Recall Complete",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            TxtDeleteStatus.Text       = $"Recall FAILED: {ex.Message}";
            TxtDeleteStatus.Foreground = System.Windows.Media.Brushes.Red;
            Log($"❌  Recall failed: {ex.Message}");
            MessageBox.Show($"Recall failed:\n\n{ex.Message}",
                "Recall Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetImportingState(false);
        }
    }

    // ─── DELETE ALL ───────────────────────────────────────────────────────────

    private async void BtnDeleteAll_Click(object sender, RoutedEventArgs e)
    {
        // Double confirmation — this is fully destructive
        var first = MessageBox.Show(
            "💣  DELETE ALL — FULL TRUNCATE\n\n" +
            "This will TRUNCATE all 10 tables written by this exporter:\n\n" +
            "  Product  ·  Stock  ·  StockTransaction\n" +
            "  Voucher  ·  AccountsTransaction  ·  PrintBarCodes\n" +
            "  ProductBrands  ·  ProductCategories  ·  ProductGroups\n" +
            "  BillNumbers\n\n" +
            "⚠️  ALL rows in these tables will be permanently removed,\n" +
            "   including data NOT imported by this tool.\n\n" +
            "Are you absolutely sure?",
            "⚠️ Confirm Delete All",
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (first != MessageBoxResult.Yes) return;

        var second = MessageBox.Show(
            "Last chance.\n\nThis will PERMANENTLY DELETE ALL DATA in those 10 tables.\nThis cannot be undone.\n\nProceed?",
            "⚠️ Final Confirmation",
            MessageBoxButton.YesNo,
            MessageBoxImage.Stop);

        if (second != MessageBoxResult.Yes) return;

        SetImportingState(true);
        TxtDeleteStatus.Text       = "Running Delete All (Truncate)...";
        TxtDeleteStatus.Foreground = System.Windows.Media.Brushes.DarkRed;
        Log("─────────────────────────────────────────────────");
        Log($"💣  DELETE ALL (TRUNCATE) started — {DateTime.Now:HH:mm:ss}");

        var progress = new Progress<(int Current, int Total, string Message)>(p =>
        {
            double pct = p.Total > 0 ? (double)p.Current / p.Total * 100 : 0;
            ImportProgress.Value  = pct;
            TxtProgressLabel.Text = $"{pct:F0}%";
            Log(p.Message);
            ScrollLog();
        });

        try
        {
            var svc     = new SqlDeleteService(TxtConnectionString.Text.Trim());
            var summary = await svc.DeleteAllAsync(progress);

            ImportProgress.Value  = 100;
            TxtProgressLabel.Text = "100%";

            TxtDeleteStatus.Text       = $"Delete All complete — {summary.TruncatedTables.Count} tables truncated  ({summary.Duration.TotalSeconds:F1}s)";
            TxtDeleteStatus.Foreground = new System.Windows.Media.SolidColorBrush(
                System.Windows.Media.Color.FromRgb(21, 128, 61));

            Log($"✅  Delete All complete — {summary.Duration.TotalSeconds:F1}s");
            Log($"    Tables truncated: {string.Join("  ·  ", summary.TruncatedTables)}");
            Log("─────────────────────────────────────────────────");

            // Reset preview — the DB is now empty so any cached rows are stale
            _previewRows            = [];
            PreviewGrid.ItemsSource = null;
            TxtRowCount.Text        = "0 rows";
            ResetBadges();

            MessageBox.Show(
                $"Delete All complete!\n\n" +
                $"Tables truncated ({summary.TruncatedTables.Count}):\n" +
                string.Join("\n", summary.TruncatedTables.Select(t => $"  • {t}")) +
                $"\n\nDuration: {summary.Duration.TotalSeconds:F1}s\n\n" +
                "The preview grid has been cleared. Re-load your Excel file before importing again.",
                "Delete All Complete",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            TxtDeleteStatus.Text       = $"Delete All FAILED: {ex.Message}";
            TxtDeleteStatus.Foreground = System.Windows.Media.Brushes.Red;
            Log($"❌  Delete All failed: {ex.Message}");
            MessageBox.Show($"Delete All failed:\n\n{ex.Message}",
                "Delete All Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
        finally
        {
            SetImportingState(false);
        }
    }

    // ─── EXPORT LOG ───────────────────────────────────────────────────────────

    private void BtnExportLog_Click(object sender, RoutedEventArgs e)
    {
        if (_lastSummary is null)
        {
            MessageBox.Show("No import has been run yet.", "No Log",
                MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dlg = new SaveFileDialog
        {
            Title    = "Save Import Log",
            Filter   = "CSV File (*.csv)|*.csv|Text File (*.txt)|*.txt",
            FileName = $"ImportLog_{DateTime.Now:yyyyMMdd_HHmmss}"
        };

        if (dlg.ShowDialog() != true) return;

        try
        {
            // C# 14: collection expression with spread
            List<string> lines =
            [
                "Row,ProductCode,ProductName,Status,Message",
                .._lastSummary.Results.Select(r =>
                    $"{r.RowNumber}," +
                    $"\"{r.ProductCode}\"," +
                    $"\"{r.ProductName.Replace("\"", "\"\"")}\"," +
                    $"{r.StatusLabel()}," +
                    $"\"{r.Message}\""),
                string.Empty,
                "── Summary ──────────────────",
                $"Voucher Mode,{_lastSummary.VoucherMode}",
                $"Total Rows,{_lastSummary.TotalRows}",
                $"New Products,{_lastSummary.NewProducts}",
                $"Stock Added,{_lastSummary.ExistingProducts}",
                $"Skipped,{_lastSummary.Skipped}",
                $"Errors,{_lastSummary.Errors}",
                $"Barcode Labels Inserted,{(_lastSummary.BarcodesEnabled ? _lastSummary.TotalBarcodesInserted.ToString() : "Disabled")}",
                ..(  _lastSummary.VoucherMode == VoucherMode.BatchSingle
                   ? (string[])[$"Batch Voucher No,{_lastSummary.BatchVoucherNo}",
                                $"Batch Total Value,{_lastSummary.BatchTotalValue:N2}"]
                   : []),
                $"Duration (s),{_lastSummary.Duration.TotalSeconds:F1}",
                $"Import Date,{_lastSummary.StartTime:yyyy-MM-dd HH:mm:ss}"
            ];

            File.WriteAllLines(dlg.FileName, lines, System.Text.Encoding.UTF8);
            Log($"📄  Log exported → {dlg.FileName}");
            MessageBox.Show("Log exported successfully.", "Export Done",
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Export failed:\n{ex.Message}", "Export Error",
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ─── HELPERS ──────────────────────────────────────────────────────────────

    private void Log(string message) =>
        Dispatcher.Invoke(() =>
        {
            ImportLog.Items.Add($"[{DateTime.Now:HH:mm:ss}]  {message}");
            ScrollLog();
        });

    private void ScrollLog()
    {
        if (ImportLog.Items.Count > 0)
            ImportLog.ScrollIntoView(ImportLog.Items[^1]);
    }

    private void SetImportingState(bool importing)
    {
        BtnImport.IsEnabled            = !importing;
        BtnImportSelected.IsEnabled    = !importing && PreviewGrid.SelectedItems.Cast<ExcelProductRow>().Any(r => r.IsValid);
        BtnCancel.IsEnabled            =  importing;
        BtnBrowse.IsEnabled            = !importing;
        RadioPerProduct.IsEnabled      = !importing;
        RadioBatchSingle.IsEnabled     = !importing;
        ChkInsertBarcodes.IsEnabled    = !importing;
        TxtConnectionString.IsReadOnly =  importing;

        // Danger Zone
        BtnDeleteAll.IsEnabled         = !importing;
        BtnRecall.IsEnabled            = !importing && _previewRows.Count > 0;
    }

    private void ResetBadges()
    {
        TxtNewCount.Text             = "✅ New: 0";
        TxtStockCount.Text           = "🔄 Stock Added: 0";
        TxtSkippedCount.Text         = "⚠️ Skipped: 0";
        TxtErrorCount.Text           = "❌ Errors: 0";
        TxtBarcodeCount.Text         = "🏷️ Labels: 0";
        BadgeBarcodeCount.Visibility = Visibility.Collapsed;
    }
}
