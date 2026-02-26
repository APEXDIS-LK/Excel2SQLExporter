using ClosedXML.Excel;
using Excel2SQLExporter.Models;

namespace Excel2SQLExporter.Services;

public class ExcelReaderService
{
    private static readonly Dictionary<string, string[]> ColumnAliases = new()
    {
        ["ProductCode"]  = ["product code", "productcode", "code", "sku", "item code"],
        ["ProductBrand"] = ["product brand", "brand", "brand name"],
        ["ProductName"]  = ["product name", "productname", "name", "description", "item name"],
        ["Category"]     = ["category", "cat", "product category", "item category"],
        ["Group"]        = ["group", "product group", "item group"],
        ["Quantity"]     = ["quantity", "qty", "stock", "opening stock", "opening qty"],
        ["CostPrice"]    = ["cost price", "cost", "purchase price", "buying price"],
        ["SellingPrice"] = ["selling price", "price", "sales price", "retail price", "jdm selling price"]
    };

    private static readonly HashSet<string> KnownSizes = new(StringComparer.OrdinalIgnoreCase)
    {
        "XS","S","M","L","XL","XXL","XXXL",
        "2XL","3XL","4XL","5XL","6XL",
        "28","29","30","31","32","33","34","35",
        "36","37","38","39","40","41","42","43","44","45","46"
    };

    public (List<ExcelProductRow> Rows, List<string> Warnings) ReadExcel(string filePath)
    {
        List<ExcelProductRow> rows     = [];
        List<string>          warnings = [];

        using var workbook  = new XLWorkbook(filePath);
        var       worksheet = workbook.Worksheets.First();

        int headerRow = FindHeaderRow(worksheet);
        if (headerRow == -1)
            throw new InvalidOperationException(
                "Could not find a header row. Ensure row 1 or 2 contains column headers.");

        var columnMap = MapColumns(worksheet, headerRow, warnings);

        foreach (var req in (string[])["ProductCode", "ProductName", "SellingPrice"])
        {
            if (!columnMap.ContainsKey(req))
                warnings.Add($"⚠️  Required column '{req}' not found. Check your Excel headers.");
        }

        int lastRow = worksheet.LastRowUsed()?.RowNumber() ?? headerRow;

        for (int r = headerRow + 1; r <= lastRow; r++)
        {
            if (worksheet.Row(r).IsEmpty()) continue;

            string productCode = GetCell(worksheet, r, columnMap, "ProductCode");
            if (string.IsNullOrWhiteSpace(productCode)) continue;

            string rawName                  = GetCell(worksheet, r, columnMap, "ProductName");
            var    (cleanName, productSize) = ExtractSize(rawName);

            var row = new ExcelProductRow
            {
                RowNumber     = r,
                ProductCode   = productCode,
                ProductBrand  = GetCell(worksheet, r, columnMap, "ProductBrand").IfEmpty("General"),
                ProductName   = rawName,
                ProductNameDb = cleanName,
                ProductSize   = productSize,
                Category      = GetCell(worksheet, r, columnMap, "Category").IfEmpty("General"),
                Group         = GetCell(worksheet, r, columnMap, "Group"),
                Quantity      = GetCell(worksheet, r, columnMap, "Quantity").ToDecimalSafe(),
                CostPrice     = GetCell(worksheet, r, columnMap, "CostPrice").ToDecimalSafe(),
                SellingPrice  = GetCell(worksheet, r, columnMap, "SellingPrice").ToDecimalSafe()
            };

            ValidateRow(row);
            rows.Add(row);
        }

        return (rows, warnings);
    }

    // ─── Size extraction ──────────────────────────────────────────────────────

    private static (string Name, string Size) ExtractSize(string rawName)
    {
        if (string.IsNullOrWhiteSpace(rawName)) return (rawName, string.Empty);

        int lastDash = rawName.LastIndexOf('-');
        if (lastDash < 0) return (rawName, string.Empty);

        string suffix = rawName[(lastDash + 1)..].Trim();

        if (KnownSizes.Contains(suffix))
            return (rawName, suffix.ToUpperInvariant());

        if (int.TryParse(suffix, out int n) && n is >= 20 and <= 60)
            return (rawName, suffix);

        return (rawName, string.Empty);
    }

    private static int FindHeaderRow(IXLWorksheet ws)
    {
        for (int r = 1; r <= 5; r++)
        {
            int textCells = 0;
            for (int c = 1; c <= 10; c++)
            {
                string val = ws.Cell(r, c).GetString();
                if (!string.IsNullOrWhiteSpace(val) && !double.TryParse(val, out _))
                    textCells++;
            }
            if (textCells >= 3) return r;
        }
        return -1;
    }

    private static Dictionary<string, int> MapColumns(
        IXLWorksheet ws, int headerRow, List<string> warnings)
    {
        Dictionary<string, int> map     = [];
        int                     lastCol = ws.Row(headerRow).LastCellUsed()?.Address.ColumnNumber ?? 20;

        for (int c = 1; c <= lastCol; c++)
        {
            string header = ws.Cell(headerRow, c).GetString().Trim().ToLowerInvariant();
            if (string.IsNullOrWhiteSpace(header)) continue;

            foreach (var (field, aliases) in ColumnAliases)
            {
                if (aliases.Any(a => a.Equals(header, StringComparison.OrdinalIgnoreCase)))
                {
                    map.TryAdd(field, c);
                    break;
                }
            }
        }

        return map;
    }

    private static string GetCell(IXLWorksheet ws, int row,
        Dictionary<string, int> map, string field)
        => map.TryGetValue(field, out int col)
            ? ws.Cell(row, col).GetString().Trim()
            : string.Empty;

    private static void ValidateRow(ExcelProductRow row)
    {
        List<string> errors = [];

        if (row.ProductCode.Length > 15)
            errors.Add($"ProductCode too long ({row.ProductCode.Length} chars, max 15)");

        if (string.IsNullOrWhiteSpace(row.ProductName))
            errors.Add("ProductName is empty");

        if (row.SellingPrice <= 0)
            errors.Add("SellingPrice is 0 or missing");

        if (errors.Count > 0)
        {
            row.IsValid           = false;
            row.ValidationMessage = string.Join("; ", errors);
        }
        else if (row.ProductName.Length > 40)
        {
            row.ValidationMessage = $"Name truncated to 40 chars (was {row.ProductName.Length})";
        }
    }
}

// ─── String helpers ───────────────────────────────────────────────────────────

public static class StringExtensions
{
    public static string IfEmpty(this string value, string fallback)
        => string.IsNullOrWhiteSpace(value) ? fallback : value;

    public static decimal ToDecimalSafe(this string value)
    {
        string clean = value.Replace(",", "").Trim();
        return decimal.TryParse(clean, out decimal result) ? result : 0m;
    }
}
