using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.OpenXmlFormats.Spreadsheet;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SpreadsheetHelper.Internal;

namespace SpreadsheetHelper;

/// <summary>
///     Represents a strongly-typed spreadsheet for working with Excel files, providing methods for reading, writing, and
///     mapping rows to instances of <typeparamref name="TEntity" />. Implements <see cref="IDisposable" /> to manage
///     workbook resources.
/// </summary>
/// <typeparam name="TEntity">
///     The type of entity to map spreadsheet rows to. Must be non-nullable and have a parameterless constructor.
/// </typeparam>
public class Spreadsheet<TEntity> : IDisposable where TEntity : new()
{
    /// <summary>
    ///     The default file name used when creating a new spreadsheet if no file name is specified.
    /// </summary>
    private const string DefaultFileName = "Book1";

    /// <summary>
    ///     The default sheet name used when creating a new spreadsheet if no sheet name is specified.
    /// </summary>
    internal const string DefaultSheetName = "Sheet1";

    /// <summary>
    ///     The array of column definitions, each representing a property of <typeparamref name="TEntity" /> mapped to a
    ///     spreadsheet column.
    /// </summary>
    private readonly ColumnInfo[] _columns;

    /// <summary>
    ///     The file path associated with the spreadsheet, used for saving and loading operations.
    /// </summary>
    private readonly string _filePath;

    /// <summary>
    ///     Indicates whether the spreadsheet contains a header row.
    /// </summary>
    private readonly bool _hasHeaderRow;

    /// <summary>
    ///     The underlying Excel workbook instance used for reading and writing spreadsheet data.
    /// </summary>
    public readonly IWorkbook Workbook;

    /// <summary>
    ///     The Excel sheet instance representing the current worksheet in the workbook.
    /// </summary>
    public readonly ISheet Sheet;


    /// <summary>
    ///     Initializes a new instance of the <see cref="Spreadsheet{TEntity}" /> class with the specified workbook, sheet
    ///     name,
    ///     and header row flag. Delegates to the private constructor, passing <c>null</c> for the file path.
    /// </summary>
    /// <param name="workbook">The Excel workbook instance to use.</param>
    /// <param name="sheetName">The name of the sheet to load from the workbook.</param>
    /// <param name="hasHeaderRow">Indicates whether the sheet contains a header row. Defaults to <c>true</c>.</param>
    public Spreadsheet(IWorkbook workbook, string sheetName, bool hasHeaderRow = true)
        : this(workbook, sheetName, hasHeaderRow, null)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="Spreadsheet{TEntity}" /> class with the specified workbook, sheet
    ///     name,
    ///     header row flag, and optional file path. Retrieves column definitions, assigns the workbook and header row flag,
    ///     and loads the specified sheet. If <paramref name="filePath" /> is null, determines the default file path based on
    ///     the workbook type; otherwise, uses the provided file path.
    /// </summary>
    /// <param name="workbook">The Excel workbook instance to use.</param>
    /// <param name="sheetName">The name of the sheet to load from the workbook.</param>
    /// <param name="hasHeaderRow">Indicates whether the sheet contains a header row.</param>
    /// <param name="filePath">
    ///     The file path associated with the spreadsheet. If null, a default path is generated based on the workbook type.
    /// </param>
    /// <exception cref="SpreadsheetException">
    ///     Thrown if the workbook type is not supported when <paramref name="filePath" /> is null.
    /// </exception>
    private Spreadsheet(IWorkbook workbook, string sheetName, bool hasHeaderRow, string? filePath)
    {
        _columns = BuildColumns();
        Workbook = workbook;
        _hasHeaderRow = hasHeaderRow;

        Sheet = Workbook.GetSheet(sheetName);
        if (Sheet is null) throw new SpreadsheetException(ResMan.Format("SpreadsheetNotFound", sheetName));

        if (filePath is null)
            _filePath = workbook switch
            {
                XSSFWorkbook => Path.Join(Environment.CurrentDirectory, $"{DefaultFileName}.xlsx"),
                HSSFWorkbook => Path.Join(Environment.CurrentDirectory, $"{DefaultFileName}.xls"),
                _ => throw new SpreadsheetException(ResMan.GetString("UnsupportedWorkbookType"))
            };
        else _filePath = filePath;
    }

    /// <summary>
    ///     Enumerates the rows of the spreadsheet, skipping the header row if present.
    ///     Yields each <see cref="IRow" /> in the sheet, allowing direct access to the underlying Excel row objects.
    /// </summary>
    public IEnumerable<IRow> Rows
    {
        get
        {
            var firstRow = _hasHeaderRow ? Sheet.FirstRowNum + 1 : Sheet.FirstRowNum;
            for (var i = firstRow; i <= Sheet.LastRowNum; i++)
            {
                var row = Sheet.GetRow(i);
                if (row is null) continue;

                // TODO: Are we relying enough on the underlying IWorkbook that changes to the row will actually be visible in Spreadsheet<TRow>.Records?
                yield return row;
            }
        }
    }

    /// <summary>
    ///     Enumerates the rows of the spreadsheet, mapping each row to an instance of <typeparamref name="TEntity" />.
    ///     Skips the header row and yields each data row as a strongly-typed entity.
    ///     After yielding, updates the row with any changes made to the entity.
    /// </summary>
    public IEnumerable<TEntity> Records
    {
        get
        {
            var firstRow = _hasHeaderRow ? Sheet.FirstRowNum + 1 : Sheet.FirstRowNum;
            for (var i = firstRow; i <= Sheet.LastRowNum; i++)
            {
                var row = Sheet.GetRow(i);
                if (row is null) continue;

                var record = GetRecord(row);
                yield return record;
                SetRow(row, record);
            }
        }
    }

    /// <summary>
    ///     Releases all resources used by the underlying workbook.
    /// </summary>
    public void Dispose()
    {
        Workbook.Dispose();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    ///     Retrieves the public properties of <typeparamref name="TEntity" /> and constructs an array of
    ///     <see cref="ColumnInfo" /> objects representing the spreadsheet columns. Each property is mapped to a column, using
    ///     the <see cref="ColumnAttribute" /> if present to determine the column name, order, and numeric format. The
    ///     resulting array is sorted by the column order.
    /// </summary>
    /// <returns>
    ///     An array of <see cref="ColumnInfo" /> objects, each representing a column mapped from a property of
    ///     <typeparamref name="TEntity" />.
    /// </returns>
    private static ColumnInfo[] BuildColumns()
    {
        // Retrieve all public properties of the TEntity type.
        var properties = typeof(TEntity).GetProperties(BindingFlags.Public | BindingFlags.Instance);

        // Initialize an array to hold column definitions for each property.
        var columns = new ColumnInfo[properties.Length];

        // Iterate through each property to build the column definitions.
        for (var i = 0; i < properties.Length; i++)
        {
            // Attempt to get the ColumnAttribute for the property, if present.
            var attribute = properties[i].GetCustomAttribute<ColumnAttribute>();
            columns[i] = new ColumnInfo
            {
                ColumnName = attribute?.Name ?? properties[i].Name,
                // Store the PropertyInfo for later use.
                Property = properties[i],
                // Use the attribute's order if specified; otherwise, use int.MaxValue.
                Order = attribute?.Order ?? int.MaxValue,
                NumericFormat = attribute?.NumericFormat
            };
        }

        // Return the columns array, ordered by the column order.
        return columns.OrderBy(c => c.Order).ToArray();
    }

    public static Spreadsheet<TEntity> Create(string path, string sheetName, bool hasHeaderRow = true)
    {
        // Determine the workbook type based on the file extension and create a new workbook instance.
        // Throws a SpreadsheetException if the file extension is not supported (.xlsx or .xls only).
        IWorkbook workbook;
        if (path.EndsWith(".xlsx")) workbook = new XSSFWorkbook();
        else if (path.EndsWith(".xls")) workbook = new HSSFWorkbook();
        else throw new SpreadsheetException("Unsupported file extension. Only .xlsx and .xls are supported.");

        // Create a new sheet in the workbook with the specified name.
        // If a header row is required, create the first row and populate it with column names.
        var sheet = workbook.CreateSheet(sheetName);
        if (!hasHeaderRow) return new Spreadsheet<TEntity>(workbook, sheetName, hasHeaderRow);

        var columns = BuildColumns();
        var row = sheet.CreateRow(0);
        for (var i = 0; i < columns.Length; i++)
        {
            var cell = row.CreateCell(i);
            cell.SetCellValue(columns[i].ColumnName);
        }

        // Instantiate a new Spreadsheet<TEntity> with the created workbook and sheet.
        return new Spreadsheet<TEntity>(workbook, sheetName, hasHeaderRow);
    }

    public static Spreadsheet<TEntity> Create(string path, string sheetName, List<TEntity> records,
        bool hasHeaderRow = true)
    {
        var spreadsheet = Create(path, sheetName, hasHeaderRow);
        foreach (var record in records) spreadsheet.AddRow(record);
        return spreadsheet;
    }

    /// <summary>
    ///     Opens an Excel file at the specified path and loads the given sheet into a <see cref="Spreadsheet{TRow}" />
    ///     instance.
    /// </summary>
    /// <param name="path">The file path of the Excel document to open.</param>
    /// <param name="sheetName">The name of the sheet to load. Defaults to <see cref="DefaultSheetName" />.</param>
    /// <param name="hasHeaderRow">Indicates whether the sheet contains a header row. Defaults to <see langword="true" />.</param>
    /// <returns>
    ///     A <see cref="Spreadsheet{TRow}" /> instance representing the specified sheet in the Excel file.
    /// </returns>
    public static Spreadsheet<TEntity> Open(string path, string sheetName = DefaultSheetName, bool hasHeaderRow = true)
    {
        // Open the Excel file for reading.
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read);
        return Open(stream, sheetName, hasHeaderRow);
    }

    /// <summary>
    ///     Asynchronously opens an Excel file at the specified path and loads the given sheet into a
    ///     <see cref="Spreadsheet{TRow}" /> instance.
    /// </summary>
    /// <param name="path">The file path of the Excel document to open.</param>
    /// <param name="sheetName">The name of the sheet to load. Defaults to <see cref="DefaultSheetName" />.</param>
    /// <param name="hasHeaderRow">Indicates whether the sheet contains a header row. Defaults to <see langword="true" />.</param>
    /// <returns>
    ///     A task representing the asynchronous operation, with a <see cref="Spreadsheet{TRow}" /> instance as the result.
    /// </returns>
    public static async Task<Spreadsheet<TEntity>> OpenAsync(string path, string sheetName = DefaultSheetName,
        bool hasHeaderRow = true)
    {
        // Open the Excel file for reading.
        await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read);
        return Open(stream, sheetName, hasHeaderRow);
    }

    /// <summary>
    ///     Opens an Excel workbook from the provided file stream and loads the specified sheet into a
    ///     <see cref="Spreadsheet{TEntity}" /> instance. Determines the workbook type based on the file extension,
    ///     and optionally specifies whether the sheet has a header row.
    /// </summary>
    /// <param name="stream">The file stream containing the Excel document.</param>
    /// <param name="sheetName">The name of the sheet to load. Defaults to <see cref="DefaultSheetName" />.</param>
    /// <param name="hasHeaderRow">Indicates whether the sheet contains a header row. Defaults to <see langword="true" />.</param>
    /// <returns>
    ///     A <see cref="Spreadsheet{TEntity}" /> instance representing the specified sheet in the Excel file.
    /// </returns>
    /// <exception cref="SpreadsheetException">
    ///     Thrown if the file extension is not supported (only .xlsx and .xls are supported).
    /// </exception>
    private static Spreadsheet<TEntity> Open(FileStream stream, string sheetName = DefaultSheetName,
        bool hasHeaderRow = true)
    {
        // Determine the workbook type based on the file extension.
        IWorkbook workbook;
        if (stream.Name.EndsWith(".xlsx")) workbook = new XSSFWorkbook(stream);
        else if (stream.Name.EndsWith(".xls")) workbook = new HSSFWorkbook(stream);
        else throw new SpreadsheetException("Unsupported file extension. Only .xlsx and .xls are supported.");

        // Create a new instance of the builder, load the specified Excel file and sheet, return the builder.
        return new Spreadsheet<TEntity>(workbook, sheetName, hasHeaderRow, stream.Name);
    }

    /// <summary>
    ///     Reads all rows from the specified Excel file and sheet, mapping each row to an instance of
    ///     <typeparamref name="TEntity" />.
    /// </summary>
    /// <param name="path">The file path of the Excel document to read.</param>
    /// <param name="sheetName">The name of the sheet to read from. Defaults to <see cref="DefaultSheetName" />.</param>
    /// <returns>
    ///     An <see cref="IReadOnlyCollection{TEntity}" /> containing all entities mapped from the sheet's rows.
    /// </returns>
    public static IReadOnlyCollection<TEntity> Read(string path, string sheetName = DefaultSheetName)
    {
        using var spreadsheet = Open(path, sheetName);
        return spreadsheet.Records.ToList();
    }

    /// <summary>
    ///     Asynchronously reads all rows from the specified Excel file and sheet, mapping each row to an instance of
    ///     <typeparamref name="TEntity" />.
    /// </summary>
    /// <param name="path">The file path of the Excel document to read.</param>
    /// <param name="sheetName">The name of the sheet to read from. Defaults to <see cref="DefaultSheetName" />.</param>
    /// <returns>
    ///     An <see cref="IAsyncEnumerable{TEntity}" /> that yields each entity mapped from the sheet's rows.
    /// </returns>
    public static async IAsyncEnumerable<TEntity> ReadAsync(string path, string sheetName = DefaultSheetName)
    {
        using var spreadsheet = await OpenAsync(path, sheetName);
        foreach (var row in spreadsheet.Records) yield return row;
    }

    /// <summary>
    ///     Saves the current spreadsheet to the original file path.
    /// </summary>
    public void Save()
    {
        SaveAs(_filePath);
    }

    /// <summary>
    ///     Asynchronously saves the current spreadsheet to the original file path.
    /// </summary>
    /// <returns>A task representing the asynchronous save operation.</returns>
    public Task SaveAsync()
    {
        return SaveAsAsync(_filePath);
    }

    /// <summary>
    ///     Saves the current spreadsheet to the specified file path.
    /// </summary>
    /// <param name="filePath">The file path to save the spreadsheet to.</param>
    public void SaveAs(string filePath)
    {
        Workbook.SaveAs(filePath);
    }

    /// <summary>
    ///     Asynchronously saves the current spreadsheet to the specified file path.
    /// </summary>
    /// <param name="filePath">The file path to save the spreadsheet to.</param>
    /// <returns>A task representing the asynchronous save operation.</returns>
    public Task SaveAsAsync(string filePath)
    {
        return Workbook.SaveAsAsync(filePath);
    }

    /// <summary>
    ///     Retrieves a row from the spreadsheet at the specified index and maps it to an instance of
    ///     <typeparamref name="TEntity" />.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the row to retrieve from the spreadsheet.</param>
    /// <returns>
    ///     An instance of <typeparamref name="TEntity" /> populated with the cell values from the specified row, or
    ///     <see langword="null" /> if the row does not exist.
    /// </returns>
    public TEntity? GetRecord(int rowIndex)
    {
        var sheetRow = Sheet.GetRow(rowIndex);
        return sheetRow is null ? default : GetRecord(sheetRow);
    }

    /// <summary>
    ///     Maps the values of the given Excel row to a new instance of <typeparamref name="TEntity" />.
    ///     Iterates through each column, retrieves the cell value, and sets the corresponding property on the entity.
    /// </summary>
    /// <param name="row">The Excel row to map to an entity.</param>
    /// <returns>
    ///     A new instance of <typeparamref name="TEntity" /> with properties populated from the row's cell values.
    /// </returns>
    private TEntity GetRecord(in IRow row)
    {
        var record = new TEntity();

        for (var i = 0; i < _columns.Length; i++)
        {
            var cell = row.Cells[i];
            var cellValue = GetCellValue(cell);
            SetPropertyValue(record, _columns[i].Property, cellValue);
        }

        return record;
    }

    /// <summary>
    ///     Adds a new row to the spreadsheet and populates it with values from the provided entity.
    /// </summary>
    /// <param name="entity">The entity whose property values will be written to the new row.</param>
    public void AddRow(in TEntity entity)
    {
        var rowIndex = Sheet.LastRowNum + 1;
        var row = Sheet.CreateRow(rowIndex);
        SetRow(row, entity);
    }

    /// <summary>
    ///     Updates the row at the specified index with values from the provided entity.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the row to update.</param>
    /// <param name="entity">The entity whose property values will be written to the row.</param>
    public void SetRow(int rowIndex, in TEntity entity)
    {
        var row = Sheet.GetRow(rowIndex);
        SetRow(row, entity);
    }

    /// <summary>
    ///     Sets the values of the given row's cells based on the properties of the provided entity. Applies any resolved cell
    ///     styles if specified in the column information.
    /// </summary>
    /// <param name="row">The Excel row to update.</param>
    /// <param name="entity">The entity whose property values will be written to the row's cells.</param>
    private void SetRow(IRow row, in TEntity entity)
    {
        for (var i = 0; i < _columns.Length; i++)
        {
            var value = _columns[i].Property.GetValue(entity);
            if (value is null) continue;

            var cell = row.GetCell(i) ?? row.CreateCell(i);
            if (_columns[i].ResolvedStyle is not null) cell.CellStyle = _columns[i].ResolvedStyle;
            SetCellValue(cell, value);
        }
    }

    /// <summary>
    ///     Removes the row at the specified index from the spreadsheet.
    /// </summary>
    /// <param name="rowIndex">The zero-based index of the row to remove.</param>
    public void RemoveRow(int rowIndex)
    {
        var row = Sheet.GetRow(rowIndex);
        RemoveRow(row);
    }

    /// <summary>
    ///     Removes the specified row from the spreadsheet.
    /// </summary>
    /// <param name="row">The Excel row to remove.</param>
    private void RemoveRow(in IRow row)
    {
        Sheet.RemoveRow(row);
    }

    /// <summary>
    ///     Gets the value from an Excel cell, handling different cell types.
    /// </summary>
    /// <param name="cell">The Excel cell to get the value from.</param>
    /// <returns>
    ///     The value of the cell, or null if the cell is null or blank. Returns a <see cref="DateTime" /> if the cell is a
    ///     date-formatted numeric cell, a <see cref="double" /> for numeric cells, a <see cref="bool" /> for boolean cells, a
    ///     <see cref="string" /> for string cells, or the result of evaluating a formula cell.
    /// </returns>
    private static object? GetCellValue(in ICell? cell)
    {
        if (cell == null) return null;

        return cell.CellType switch
        {
            CellType.Numeric when DateUtil.IsCellDateFormatted(cell) => cell.DateCellValue,
            CellType.Numeric => cell.NumericCellValue,
            CellType.Boolean => cell.BooleanCellValue,
            CellType.String => cell.StringCellValue,
            CellType.Formula => TryGetNumericValue(cell),
            CellType.Blank => null,
            CellType.Error => null,
            CellType.Unknown => null,
            _ => cell.StringCellValue
        };

        // Helper function to try getting the numeric value from a formula cell.
        object TryGetNumericValue(ICell numericCell)
        {
            try
            {
                return numericCell.NumericCellValue;
            }

            catch (FormatException)
            {
                return numericCell.StringCellValue;
            }
        }
    }

    /// <summary>
    ///     Sets the value of the specified Excel cell based on the provided value's type. Handles various native and
    ///     additional types, mapping them to the appropriate Excel cell type and value. For unsupported types, attempts to
    ///     convert the value to <see cref="double" />.
    /// </summary>
    /// <param name="cell">The Excel cell to set the value for.</param>
    /// <param name="value">
    ///     The value to assign to the cell. Supported types include <see cref="DateOnly" />, <see cref="DateTime" />,
    ///     <see cref="XSSFRichTextString" />, <see cref="HSSFRichTextString" />, <see cref="bool" />, <see cref="double" />,
    ///     <see cref="string" />, <see cref="DateTimeOffset" />, <see cref="byte" />, <see cref="sbyte" />,
    ///     <see cref="char" />, <see cref="decimal" />, <see cref="float" />, <see cref="int" />, <see cref="uint" />,
    ///     <see cref="nint" />, <see cref="nuint" />, <see cref="long" />, <see cref="ulong" />, <see cref="short" />, and
    ///     <see cref="ushort" />.
    /// </param>
    private static void SetCellValue(ICell cell, object? value)
    {
        // Write the type mapped object to the cell.
        switch (value)
        {
            // START: Native types
            case DateOnly typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(typedValue);
                break;

            case DateTime typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(typedValue);
                break;

            case XSSFRichTextString typedValue:
                cell.SetCellType(CellType.String);
                cell.SetCellValue(typedValue);
                break;

            case HSSFRichTextString typedValue:
                cell.SetCellType(CellType.String);
                cell.SetCellValue(typedValue);
                break;

            case bool typedValue:
                cell.SetCellType(CellType.Boolean);
                cell.SetCellValue(typedValue);
                break;

            case double typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(typedValue);
                break;

            case string typedValue:
                cell.SetCellType(CellType.String);
                cell.SetCellValue(typedValue);
                break;
            // END: Natively supported types

            // START: Additional types
            case DateTimeOffset typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(typedValue.DateTime);
                break;

            case byte typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(typedValue);
                break;

            case sbyte typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(typedValue);
                break;

            case char typedValue:
                cell.SetCellType(CellType.String);
                cell.SetCellValue(typedValue);
                break;

            case decimal typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(Convert.ToDouble(typedValue));
                break;

            case float typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(typedValue);
                break;

            case int typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(typedValue);
                break;

            case uint typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(typedValue);
                break;

            case nint typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(typedValue);
                break;

            case nuint typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(typedValue);
                break;

            case long typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(typedValue);
                break;

            case ulong typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(typedValue);
                break;

            case short typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(typedValue);
                break;

            case ushort typedValue:
                cell.SetCellType(CellType.Numeric);
                cell.SetCellValue(typedValue);
                break;

            default:
                cell.SetCellType(CellType.Unknown);
                cell.SetCellValue(Convert.ToDouble(value));
                break;
            // END: Additional types
        }
    }

    /// <summary>
    ///     Sets the value of the specified property on the given <typeparamref name="TEntity" /> instance, converting the
    ///     provided value to the property's type if necessary. Handles nullable property types and throws a
    ///     <see cref="SpreadsheetException" /> if the conversion fails.
    /// </summary>
    /// <param name="row">The instance of <typeparamref name="TEntity" /> whose property value is to be set.</param>
    /// <param name="property">The <see cref="PropertyInfo" /> representing the property to set.</param>
    /// <param name="value">The value to assign to the property, which will be converted to the property's type if needed.</param>
    /// <exception cref="SpreadsheetException">
    ///     Thrown when the value cannot be converted to the property's type due to a format or cast error.
    /// </exception>
    private static void SetPropertyValue(TEntity row, PropertyInfo property, object? value)
    {
        var propertyType = property.PropertyType;
        if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
        {
            var underlyingType = Nullable.GetUnderlyingType(propertyType);
            if (underlyingType is not null) propertyType = underlyingType;
        }

        try
        {
            property.SetValue(row, Convert.ChangeType(value, propertyType));
        }

        catch (FormatException ex)
        {
            var valueType = value?.GetType().Name ?? "null";
            throw new SpreadsheetException(ResMan.Format("ConversionFailed",
                valueType, property.Name, property.PropertyType.Name), ex);
        }

        catch (InvalidCastException ex)
        {
            var valueType = value?.GetType().Name ?? "null";
            throw new SpreadsheetException(ResMan.Format("InvalidCast",
                valueType, property.Name, property.PropertyType.Name), ex);
        }
    }
}