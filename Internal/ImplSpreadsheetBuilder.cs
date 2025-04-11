// A .NET library to simplify reading and writing Excel files using NPOI.
// Copyright (C) 2025 Will Branch
//
// This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public
// License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later
// version.
//
// This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied
// warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License along with this program. If not, see
// <https://www.gnu.org/licenses/>.

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelHelper.Internal;

/// <summary>
///     Base class for spreadsheet builders, providing core functionality for managing the workbook and worksheet.
/// </summary>
public class ImplSpreadsheetBuilder : IDisposable
{
    internal const string SheetName = "Sheet1";

    /// <summary>
    ///     A list to store the definitions of the columns in the Excel sheet.
    /// </summary>
    internal readonly List<ColumnInfo> ColumnDefinitions = [];

    /// <summary>
    ///     Initializes a new instance of the <see cref="ImplSpreadsheetBuilder" /> class, creating a new
    ///     <see cref="IWorkbook" /> with a default sheet name.
    /// </summary>
    protected ImplSpreadsheetBuilder() : this(new XSSFWorkbook())
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ImplSpreadsheetBuilder" /> class with a specified sheet name, creating
    ///     a new <see cref="IWorkbook" />.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to create.</param>
    protected ImplSpreadsheetBuilder(string sheetName) : this(new XSSFWorkbook(), sheetName)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ImplSpreadsheetBuilder" /> class with an existing
    ///     <see cref="IWorkbook" />
    ///     and an optional sheet name.
    /// </summary>
    /// <param name="workbook">The existing Excel workbook to use.</param>
    /// <param name="sheetName">The name of the sheet to work with. Defaults to "Sheet1".</param>
    protected ImplSpreadsheetBuilder(IWorkbook workbook, string sheetName = SheetName)
    {
        Workbook = workbook;
        Worksheet = Workbook.GetSheet(sheetName) ?? workbook.CreateSheet(sheetName);
    }

    /// <summary>Gets the underlying Excel workbook.</summary>
    public IWorkbook Workbook { get; private set; }

    /// <summary>Gets the current worksheet being built.</summary>
    public ISheet Worksheet { get; private set; }

    /// <inheritdoc />
    public void Dispose()
    {
        Workbook.Dispose();
        GC.SuppressFinalize(this);
    }

    /// <summary>
    ///     Sets the style for a specific column by its name.
    /// </summary>
    /// <param name="columnName">The name of the column to style.</param>
    /// <param name="configureStyle">An action to configure the cell style using the <see cref="CellStyleBuilder" />.</param>
    protected void SetColumnStyle(string columnName, Action<CellStyleBuilder> configureStyle)
    {
        var column = ColumnDefinitions.FirstOrDefault(x => x.ColumnName == columnName);
        if (column is null)
        {
            column = new ColumnInfo(columnName)
            {
                StyleConfiguration = style =>
                {
                    var styleBuilder = new CellStyleBuilder(Workbook, style);
                    configureStyle(styleBuilder);
                }
            };

            ColumnDefinitions.Add(column);
        }
        else
        {
            column.StyleConfiguration = style =>
            {
                var styleBuilder = new CellStyleBuilder(Workbook, style);
                configureStyle(styleBuilder);
            };
        }
    }

    /// <summary>
    ///     Sets the numeric format for a specific column by its name.
    /// </summary>
    /// <param name="columnName">The name of the column to format.</param>
    /// <param name="numericFormat">The numeric format string to apply (e.g., "0.00", "#,##0").</param>
    protected void SetColumnFormat(string columnName, string numericFormat)
    {
        var column = ColumnDefinitions.FirstOrDefault(x => x.ColumnName == columnName);
        if (column is null)
        {
            column = new ColumnInfo(columnName) { NumericFormat = numericFormat };
            ColumnDefinitions.Add(column);
        }
        else
        {
            column.NumericFormat = numericFormat;
        }
    }

    /// <summary>
    ///     Sets the value of an Excel cell based on the type of the provided object.
    /// </summary>
    /// <param name="cell">The Excel cell to set the value for.</param>
    /// <param name="value">
    ///     The object value to write to the cell. Supports various primitive types, DateTime,
    ///     DateTimeOffset, and RichTextString.
    /// </param>
    internal static void SetValue(ref ICell cell, object value)
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
    ///     Builds the column definition, applying style and format if specified.
    /// </summary>
    /// <param name="column">The column definition to build.</param>
    /// <returns>The built column definition.</returns>
    internal ColumnInfo BuildColumn(ColumnInfo column)
    {
        if (column.StyleConfiguration is not null)
        {
            var style = Workbook.CreateCellStyle();
            column.StyleConfiguration(style);

            if (column.NumericFormat is not null)
            {
                var format = Workbook.CreateDataFormat();
                style.DataFormat = format.GetFormat(column.NumericFormat);
            }

            column.ResolvedStyle = style;
        }

        if (column.NumericFormat is null) return column;
        {
            if (column.ResolvedStyle is null)
            {
                var style = Workbook.CreateCellStyle();
                column.ResolvedStyle = style;
            }

            var format = Workbook.CreateDataFormat();
            column.ResolvedStyle.DataFormat = format.GetFormat(column.NumericFormat);
        }

        return column;
    }

    internal void ReorderColumns(List<ColumnInfo> columns)
    {
        // Order the columns based on the Order property of the ExcelColumnAttribute then add columns without a
        // specified order to the end of the list.
        var ordered = columns.Where(c => c.Order > -1).OrderBy(c => c.Order).ToList();
        ordered.AddRange(columns.Where(c => c.Order == -1));

        ColumnDefinitions.Clear();
        ColumnDefinitions.AddRange(ordered);
    }
}