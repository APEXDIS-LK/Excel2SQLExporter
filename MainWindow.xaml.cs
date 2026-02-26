using System.Collections.ObjectModel;
using System.Globalization;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Media;
using ExcelToSqlImporter.Models;
using ExcelToSqlImporter.Services;
using Microsoft.Win32;

namespace ExcelToSqlImporter;

// ─── C# 14: primary constructor — no boilerplate constructor body ─────────────
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
    // C# 14: field keyword — lazy-initialised, no separate backing field
    private ExcelReaderService ExcelReader
    {
        get => field ??= new ExcelReaderService();
    }

    private List<ExcelProductRow> _previewRows = [];       // C# 14: collection expression
    private CancellationTokenSource? _cts;
    private ImportSummary? _lastSummary;

    public MainWindow()
    {
        InitializeComponent();
        Log("👋  Ready. Paste your connection string → select an Excel file → Preview → Import.");
        Log("ℹ️   Each row writes to 9 tables: Brands, Categories, Groups, Product, Stock,");
        Log("     BillNumbers, StockTransaction, Voucher, AccountsTransaction.");
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

        // Reset state
        _previewRows            = [];
        PreviewGrid.ItemsSource = null;
        TxtRowCount.Text        = "0 rows";
        BtnImport.IsEnabled     = false;
        BadgeWarnings.Visibility = Visibility.Collapsed;
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

            Log($"✅  Preview loaded — {rows.Count} rows total:");
            Log($"     ✅ Valid          : {validCount}");
            Log($"     ⚠️  Invalid/skipped: {invalidCount}");
            Log($"     📐 With size      : {withSize}  (ProductSize column in Stock)");
            Log($"     ✂️  Name truncated : {truncated} (>40 chars, truncated for Product table)");

            if (invalidCount > 0)
                Log($"⚠️  Hover the ⚠️ icon in the grid for row-level details.");

            BtnImport.IsEnabled = validCount > 0;
        }
        catch (Exception ex)
        {
            Log($"❌  Failed to read file: {ex.Message}");
            MessageBox.Show($"Could not read the Excel file:\n\n{ex.Message}",
                "Read Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    // ─── IMPORT ───────────────────────────────────────────────────────────────

    private async void BtnImport_Click(object sender, RoutedEventArgs e)
    {
        if (_previewRows.Count == 0)
        {
            MessageBox.Show("No data to import. Please preview the file first.",
                "No Data", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        string connStr = TxtConnectionString.Text.Trim();
        if (string.IsNullOrWhiteSpace(connStr))
        {
            MessageBox.Show("Please enter a SQL Server connection string.",
                "No Connection", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        int validCount = _previewRows.Count(r => r.IsValid);

        var confirm = MessageBox.Show(
            $"Ready to import {validCount} valid rows into SimplePOSDB.\n\n" +
            "For each row the tool will:\n" +
            "  1.  Check / insert  ProductBrands\n" +
            "  2.  Check / insert  ProductCategories\n" +
            "  3.  Check / insert  ProductGroups  (default: General)\n" +
            "  4.  Insert          Product         (only if new ProductCode)\n" +
            "  5.  Insert          Stock           (always — with ProductSize)\n" +
            "  6.  Increment       BillNumbers.JournalVoucherNo\n" +
            "  7.  Insert          StockTransaction\n" +
            "  8.  Insert          Voucher         (Opening Stock Balance)\n" +
            "  9.  Insert 2 rows   AccountsTransaction  (Debit 141 / Credit 92)\n\n" +
            "Each row is its own transaction — one failure won't affect others.\n\n" +
            "Continue?",
            "Confirm Import",
            MessageBoxButton.YesNo,
            MessageBoxImage.Question);

        if (confirm != MessageBoxResult.Yes) return;

        // ── Set up UI for import ──────────────────────────────────────────────
        SetImportingState(true);
        ResetBadges();
        ImportLog.Items.Clear();
        ImportProgress.Value  = 0;
        TxtProgressLabel.Text = "0%";
        Log($"🚀  Import started — {DateTime.Now:HH:mm:ss}");
        Log($"    {validCount} rows to process...");

        _cts = new CancellationTokenSource();

        var progress = new Progress<(int Current, int Total, string Message)>(p =>
        {
            double pct            = (double)p.Current / p.Total * 100;
            ImportProgress.Value  = pct;
            TxtProgressLabel.Text = $"{pct:F0}%";
            Log(p.Message);
            ScrollLog();
        });

        try
        {
            _lastSummary = await new SqlImportService(connStr)
                .ImportAsync(_previewRows, progress, _cts.Token);

            // ── Update badges ─────────────────────────────────────────────────
            TxtNewCount.Text     = $"✅ New: {_lastSummary.NewProducts}";
            TxtStockCount.Text   = $"🔄 Stock Added: {_lastSummary.ExistingProducts}";
            TxtSkippedCount.Text = $"⚠️ Skipped: {_lastSummary.Skipped}";
            TxtErrorCount.Text   = $"❌ Errors: {_lastSummary.Errors}";
            ImportProgress.Value  = 100;
            TxtProgressLabel.Text = "100%";

            // ── Summary log ───────────────────────────────────────────────────
            Log("─────────────────────────────────────────────────");
            Log($"✅  Import complete — {_lastSummary.EndTime:HH:mm:ss}");
            Log($"    Total rows    : {_lastSummary.TotalRows}");
            Log($"    New products  : {_lastSummary.NewProducts}");
            Log($"    Stock added   : {_lastSummary.ExistingProducts}");
            Log($"    Skipped       : {_lastSummary.Skipped}");
            Log($"    Errors        : {_lastSummary.Errors}");
            Log($"    Duration      : {_lastSummary.Duration.TotalSeconds:F1}s");
            Log("─────────────────────────────────────────────────");

            // ── Show errors in log ────────────────────────────────────────────
            if (_lastSummary.Errors > 0)
            {
                Log("❌  Rows with errors:");
                foreach (var r in _lastSummary.Results.Where(r => r.Status == ImportStatus.Error))
                    Log($"     Row {r.RowNumber} [{r.ProductCode}]: {r.Message}");
            }

            MessageBox.Show(
                $"Import complete!\n\n" +
                $"✅  New products  : {_lastSummary.NewProducts}\n" +
                $"🔄  Stock added   : {_lastSummary.ExistingProducts}\n" +
                $"⚠️   Skipped       : {_lastSummary.Skipped}\n" +
                $"❌  Errors        : {_lastSummary.Errors}\n\n" +
                $"Duration: {_lastSummary.Duration.TotalSeconds:F1} seconds\n\n" +
                "Use 'Export Log as CSV' to save a detailed report.",
                "Import Complete",
                MessageBoxButton.OK,
                _lastSummary.Errors > 0 ? MessageBoxImage.Warning : MessageBoxImage.Information);
        }
        catch (OperationCanceledException)
        {
            Log("⛔  Import cancelled by user.");
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
            // C# 14: collection expression with spread operator
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
                "── Summary ──",
                $"Total Rows,{_lastSummary.TotalRows}",
                $"New Products,{_lastSummary.NewProducts}",
                $"Stock Added,{_lastSummary.ExistingProducts}",
                $"Skipped,{_lastSummary.Skipped}",
                $"Errors,{_lastSummary.Errors}",
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
        BtnCancel.IsEnabled            =  importing;
        BtnBrowse.IsEnabled            = !importing;     // direct x:Name reference — no FindName needed
        TxtConnectionString.IsReadOnly =  importing;
    }

    private void ResetBadges()
    {
        TxtNewCount.Text     = "✅ New: 0";
        TxtStockCount.Text   = "🔄 Stock Added: 0";
        TxtSkippedCount.Text = "⚠️ Skipped: 0";
        TxtErrorCount.Text   = "❌ Errors: 0";
    }
}
