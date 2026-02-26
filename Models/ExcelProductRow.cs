namespace Excel2SQLExporter.Models;

/// <summary>
/// One row read from the Excel file.
/// C# 14: 'field' keyword enforces trim / truncate / default rules on set
/// without declaring separate private backing fields.
/// </summary>
public class ExcelProductRow
{
    public int RowNumber { get; set; }

    public string ProductCode
    {
        get;
        set => field = value.Trim();
    } = string.Empty;

    public string ProductBrand
    {
        get;
        set => field = value.Trim();
    } = string.Empty;

    /// Full original name as read from Excel — shown in preview grid
    public string ProductName
    {
        get;
        set => field = value.Trim();
    } = string.Empty;

    /// Truncated to 40 chars — written to Product.ProductName in SQL
    public string ProductNameDb
    {
        get;
        set => field = value.Length > 40 ? value[..40] : value;
    } = string.Empty;

    /// Size extracted from after the last '-' in ProductName (e.g. "XXL", "36")
    /// Written to Stock.ProductSize; NULL in DB if empty
    public string ProductSize
    {
        get;
        set => field = value.Trim().ToUpperInvariant();
    } = string.Empty;

    public string Category
    {
        get;
        set => field = string.IsNullOrWhiteSpace(value) ? "General" : value.Trim();
    } = "General";

    public string Group
    {
        get;
        set => field = string.IsNullOrWhiteSpace(value) ? "General" : value.Trim();
    } = "General";

    public decimal Quantity     { get; set; }
    public decimal CostPrice    { get; set; }
    public decimal SellingPrice { get; set; }

    public bool   IsValid           { get; set; } = true;
    public string ValidationMessage { get; set; } = string.Empty;

    /// Qty × CostPrice — used for Voucher and AccountsTransaction amounts
    public decimal TotalValue => Quantity * CostPrice;
}
