# Excel → SQL Server Importer
### SimplePOSDB · Opening Stock Entry

**Stack:** .NET 10 · C# 14 · WPF · Visual Studio 2026
**Platform:** Windows 10 / 11 Pro (x64)

---

## What It Does

Imports product opening stock from `.xlsx` / `.xlsm` into **SimplePOSDB**.
Every valid Excel row writes to **9 tables** in a single atomic transaction:

| Step | Table | Action |
|------|-------|--------|
| 1 | `ProductBrands` | Check name → insert if new → return `BrandId` |
| 2 | `ProductCategories` | Check name → insert if new → return `CategoryId` |
| 3 | `ProductGroups` | Check name → insert if new → return `GroupId` (default: **General**) |
| 4 | `Product` | Insert **only if ProductCode not found** |
| 5 | `Stock` | Always insert — includes **ProductSize** extracted from product name |
| 6 | `BillNumbers` | Read `JournalVoucherNo` → increment → update |
| 7 | `StockTransaction` | Always insert — uses real `JUR` voucher number from step 6 |
| 8 | `Voucher` | Insert — Type: `Opening Stock Balance`, Debit: `141`, Credit: `92` |
| 9 | `AccountsTransaction` | Insert **2 rows** — DEBIT `141` (Stock), CREDIT `92` (Opening Balance Equity) |

**Each row is its own transaction.** One failure rolls back only that row.

---

## ProductSize Extraction

Sizes are extracted automatically from the product name (after the last `-`):

```
"POLO BIG SIZE - XXL"      →  ProductSize = "XXL"
"POLO STRIPES - M"         →  ProductSize = "M"
"BAGGY DENIM PANT BK- 36"  →  ProductSize = "36"
"POLO SHIRT"               →  ProductSize = NULL
```

**Clothing sizes recognised:** `XS S M L XL XXL XXXL 2XL 3XL 4XL 5XL 6XL`
**Numeric sizes recognised:** `28 – 46` (waist / chest)
Anything else after a `-` is ignored — ProductSize stays NULL.

---

## Excel Column Headers (flexible, case-insensitive)

| Field | Accepted Headers |
|-------|-----------------|
| Product Code | `Product Code`, `ProductCode`, `Code`, `SKU` |
| Brand | `Product Brand`, `Brand`, `Brand Name` |
| Product Name | `Product Name`, `Name`, `Description` |
| Category | `Category`, `Product Category` |
| **Group** | `Group`, `Product Group` — **optional, defaults to "General"** |
| Quantity | `Quantity`, `Qty`, `Opening Stock` |
| Cost Price | `Cost Price`, `Cost`, `Purchase Price` |
| Selling Price | `Selling Price`, `Price`, `JDM Selling Price`, `Sales Price` |

---

## Journal Voucher Numbering

- Format: `JUR` + 7 digits → `JUR0000001`
- If `BillNumbers` table has no row → creates one, starts at `JUR0000001`
- If row exists → reads `JournalVoucherNo`, increments, updates, uses new number
- **One unique journal number per product row**

---

## C# 14 Features Used

| Feature | File | Usage |
|---------|------|-------|
| `field` keyword (semi-auto properties) | `ExcelProductRow.cs` | Trim, truncate, default on set — no backing fields |
| Primary constructor | `SqlImportService.cs` | `public class SqlImportService(string connectionString)` |
| Collection expressions `[]` | Throughout | `List<ExcelProductRow> rows = []` |
| Collection spread `..` | `MainWindow.xaml.cs` | `List<string> lines = ["header", .._results.Select(...)]` |
| Raw string literals `"""` | `SqlImportService.cs` | All SQL INSERT statements |
| Lazy `field` initialisation | `MainWindow.xaml.cs` | `ExcelReaderService` property |
| Pattern matching `is >= and <=` | `ExcelReaderService.cs` | `n is >= 20 and <= 60` |

---

## Prerequisites

- **Visual Studio 2026** with **.NET Desktop Development** workload
- **.NET 10 SDK** — https://dotnet.microsoft.com/download/dotnet/10.0
- **SQL Server Express** with `SimplePOSDB` database created from your SQL script

---

## Quick Start

1. Extract ZIP
2. Open `ExcelToSqlImporter.csproj` in Visual Studio 2026
3. NuGet auto-restores `ClosedXML` and `Microsoft.Data.SqlClient` on first build
4. Press **F5**

---

## Connection String Examples

```
# Windows Auth — most common with SQL Server Express
Server=.\SQLEXPRESS;Database=SimplePOSDB;Trusted_Connection=True;TrustServerCertificate=True;

# SQL Server Auth
Server=.\SQLEXPRESS;Database=SimplePOSDB;User Id=sa;Password=YourPassword;TrustServerCertificate=True;

# Named PC
Server=OFFICE-PC\SQLEXPRESS;Database=SimplePOSDB;Trusted_Connection=True;TrustServerCertificate=True;
```

---

## Project Structure

```
ExcelToSqlImporter/
├── ExcelToSqlImporter.csproj      .NET 10, LangVersion 14
├── App.xaml                       App styles — buttons, cards, textboxes
├── App.xaml.cs                    WPF app entry point with InitializeComponent
├── MainWindow.xaml                Full UI — all x:Name bindings verified
├── MainWindow.xaml.cs             All event handlers — C# 14 features
├── Models/
│   ├── ExcelProductRow.cs         field keyword properties + TotalValue
│   └── ImportResult.cs            ImportStatus enum + static extension methods
└── Services/
    ├── ExcelReaderService.cs      Header detection, size extraction, validation
    └── SqlImportService.cs        Primary constructor, 9-table atomic import
```

---

## NuGet Packages

| Package | Version | Purpose |
|---------|---------|---------|
| `ClosedXML` | 0.104.1 | Read `.xlsx` / `.xlsm` files |
| `Microsoft.Data.SqlClient` | 5.2.2 | SQL Server connection + commands |
