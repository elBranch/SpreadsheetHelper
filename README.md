# ExcelHelper

A .NET library to simplify reading and writing Excel files using NPOI.

## Overview

ExcelHelper provides a streamlined way to interact with Excel files (.xlsx and .xls) in your .NET applications. It leverages the NPOI library to handle the complexities of the Excel file format, allowing you to focus on your data.

Key Features:

* **Read Excel Data:** Easily read Excel data into `DataTable` objects or lists of custom classes.
* **Write Excel Data:** Generate Excel files from `DataTable` objects or lists of custom classes.
* **Column Mapping:** Map Excel columns to class properties using attributes.
* **Styling:** Apply cell styles (fonts, colors, borders, alignment) to customize the appearance of your Excel files.
* **Asynchronous Operations:** Supports asynchronous file I/O for improved performance.
* **Fluent API:** Provides a fluent interface for building Excel spreadsheets.

## Usage
### Writing Excel Files
#### From a List of Objects
```csharp
using ExcelHelper;

public class MyData
{
    [Column("Name", Order = 1)]
    public string Name { get; set; }

    [Column("Age", Order = 2)]
    public int Age { get; set; }

    [Column("Date of Birth", Order = 3, NumericFormat = "yyyy-MM-dd")]
    public DateTime? DateOfBirth { get; set; }
}

// ...

var data = new List<MyData>
{
    new MyData { Name = "Alice", Age = 30, DateOfBirth = new DateTime(1993, 1, 1) },
    new MyData { Name = "Bob", Age = 25, DateOfBirth = new DateTime(1998, 5, 10) }
};

using var builder = new SpreadsheetBuilder<MyData>();
builder.SetColumnStyle(x => x.Name, style => style.SetBold().Alignment(NPOI.SS.UserModel.HorizontalAlignment.Center))
    .SetColumnStyle(x => x.Age, style => style.Alignment(NPOI.SS.UserModel.HorizontalAlignment.Right))
    .Build(data)
    .SaveAs("output.xlsx");
```

#### From a `DataTable`
```csharp
using ExcelHelper;
using System.Data;

// ...

var dataTable = new DataTable();
dataTable.Columns.Add("Product", typeof(string));
dataTable.Columns.Add("Price", typeof(decimal));
dataTable.Rows.Add("Laptop", 1200.00m);
dataTable.Rows.Add("Mouse", 25.00m);

using var builder = new SpreadsheetBuilder();
builder.SetColumnStyle("Product", style => style.SetBold())
    .SetColumnFormat("Price", "$#,##0.00")
    .Build(dataTable)
    .SaveAs("products.xlsx");
```

### Reading Excel Files
#### To a `DataTable`
```csharp
using ExcelHelper;
using System.Data;

// ...

var dataTable = await SpreadsheetReader.ReadFileAsync("data.xlsx");

foreach (DataRow row in dataTable.Rows)
{
    Console.WriteLine(string.Join(", ", row.ItemArray));
}
```

#### To a List of Objects
```csharp
using ExcelHelper;

public class MyData
{
    public string Name { get; set; }
    public int Age { get; set; }
    public DateTime? DateOfBirth { get; set; }
}

// ...

var data = await SpreadsheetReader.ReadFileAsync<MyData>("data.xlsx");

foreach (var item in data)
{
    Console.WriteLine($"Name: {item.Name}, Age: {item.Age}, DOB: {item.DateOfBirth}");
}
```

## API Reference
### `SpreadsheetBuilder<TRow>`

* `Open(string path, string sheetName = "Sheet1")`
  * Opens an existing Excel file for writing.
* `OpenAsync(string path, string sheetName = "Sheet1")`
  * Asynchronously opens an existing Excel file.
* `SetColumnStyle<TColumn>(Expression<Func<TRow, TColumn>> propertyExpression, Action<CellStyleBuilder> configureStyle)`
  * Sets the style for a column.
* `Build(IList<TRow> records)`
  * Builds the Excel workbook from a list of records.

### `SpreadsheetBuilder`

* `Open(string path, string sheetName = "Sheet1")`
  * Opens an existing Excel file for writing (for DataTable).
* `OpenAsync(string path, string sheetName = "Sheet1")`
  * Asynchronously opens an existing Excel file (for DataTable).
* `SetColumnStyle(string columnName, Action<CellStyleBuilder> configureStyle)`
  * Sets the style for a column (by name).
* `SetColumnFormat(string columnName, string numericFormat)`
  * Sets the numeric format for a column.
* `Build(DataTable dataTable)`
  * Builds the Excel workbook from a DataTable.

### `SpreadsheetReader`

* `ReadFileAsync(string filePath, string? sheetName = "Sheet1")`
  * Asynchronously reads an Excel file into a DataTable.
* `ReadFileAsync<TRow>(string filePath, string? sheetName = "Sheet1") where TRow : new()`
  * Asynchronously reads an Excel file into a list of objects.

### `WorkbookHelper`

* `SaveAs(this IWorkbook workbook, string path)`
  * Saves the workbook synchronously.
* `SaveAsAsync(this IWorkbook workbook, string path)`
  * Saves the workbook asynchronously.

### `ColumnAttribute`
Attribute to customize column name and order.

### `CellStyleBuilder`
Fluent API for configuring cell styles.