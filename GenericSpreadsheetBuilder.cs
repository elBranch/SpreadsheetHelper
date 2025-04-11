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

using System.Linq.Expressions;
using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SpreadsheetHelper.Internal;

namespace SpreadsheetHelper;

/// <summary>
///     Builds an Excel spreadsheet for a specific row type <typeparamref name="TRow" />.
/// </summary>
/// <typeparam name="TRow">The type representing a single row in the spreadsheet.</typeparam>
public class SpreadsheetBuilder<TRow> : ImplSpreadsheetBuilder
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="SpreadsheetBuilder{TRow}" /> class with the default sheet name of
    ///     Sheet1, creating a new <see cref="IWorkbook" />.
    /// </summary>
    public SpreadsheetBuilder()
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="SpreadsheetBuilder{TRow}" /> class with a specified sheet name,
    ///     creating a new <see cref="IWorkbook" />.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to create.</param>
    public SpreadsheetBuilder(string sheetName) : base(sheetName)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="ImplSpreadsheetBuilder" /> class with an existing
    ///     <see cref="IWorkbook" /> and an optional sheet name.
    /// </summary>
    /// <param name="workbook">The existing Excel workbook to use.</param>
    /// <param name="sheetName">The name of the sheet to work with. Defaults to "Sheet1".</param>
    public SpreadsheetBuilder(IWorkbook workbook, string sheetName = SheetName) : base(workbook, sheetName)
    {
    }

    /// <summary>
    ///     Opens an existing Excel file and initializes a <see cref="SpreadsheetBuilder{TRow}" /> for writing to a specific
    ///     sheet.
    /// </summary>
    /// <param name="path">The path to the Excel file.</param>
    /// <param name="sheetName">The name of the sheet to work with. Defaults to "Sheet1".</param>
    /// <returns>A new instance of <see cref="SpreadsheetBuilder{TRow}" />.</returns>
    public static SpreadsheetBuilder<TRow> Open(string path, string sheetName = SheetName)
    {
        // Open the Excel file for reading.
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read);

        // Determine the workbook type based on the file extension.
        IWorkbook workbook;
        if (path.EndsWith(".xlsx")) workbook = new XSSFWorkbook(stream);
        else if (path.EndsWith(".xls")) workbook = new HSSFWorkbook(stream);
        else throw new SpreadsheetException("Unsupported file extension. Only .xlsx and .xls are supported.");

        // Create a new instance of the builder, load the specified Excel file and sheet, return the builder.
        return new SpreadsheetBuilder<TRow>(workbook, sheetName);
    }

    /// <summary>
    ///     Asynchronously opens an existing Excel file and initializes a <see cref="SpreadsheetBuilder{TRow}" /> for writing
    ///     to a specific sheet.
    /// </summary>
    /// <param name="path">The path to the Excel file.</param>
    /// <param name="sheetName">The name of the sheet to work with. Defaults to "Sheet1".</param>
    /// <returns>A new instance of <see cref="SpreadsheetBuilder{TRow}" />.</returns>
    public static async Task<SpreadsheetBuilder<TRow>> OpenAsync(string path, string sheetName = SheetName)
    {
        // Open the Excel file for reading.
        await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read);

        // Determine the workbook type based on the file extension.
        IWorkbook workbook;
        if (path.EndsWith(".xlsx")) workbook = new XSSFWorkbook(stream);
        else if (path.EndsWith(".xls")) workbook = new HSSFWorkbook(stream);
        else throw new SpreadsheetException("Unsupported file extension. Only .xlsx and .xls are supported.");

        // Create a new instance of the builder, load the specified Excel file and sheet, return the builder.
        return new SpreadsheetBuilder<TRow>(workbook, sheetName);
    }

    /// <summary>
    ///     Sets the style for a specific column based on the property of the row type.
    /// </summary>
    /// <typeparam name="TColumn">The property type of the column.</typeparam>
    /// <param name="propertyExpression">An expression selecting the property to style (e.g., row => row.PropertyName).</param>
    /// <param name="configureStyle">An action to configure the cell style using the <see cref="CellStyleBuilder" />.</param>
    /// <returns>The current <see cref="SpreadsheetBuilder{TRow}" /> instance for method chaining.</returns>
    /// <exception cref="ArgumentException">Thrown if the provided expression does not select a property.</exception>
    public SpreadsheetBuilder<TRow> SetColumnStyle<TColumn>(Expression<Func<TRow, TColumn>> propertyExpression,
        Action<CellStyleBuilder> configureStyle)
    {
        // Check if the provided expression is a property access, throwing if not.
        if (propertyExpression.Body is not MemberExpression { Member: PropertyInfo propertyInfo })
            throw new SpreadsheetException($"The column specified ({propertyExpression.Name}) " +
                                           $"is not a property of the row.");

        // Call the base class's SetColumnStyle method with the property name.
        base.SetColumnStyle(propertyInfo.Name, configureStyle);
        return this;
    }

    /// <summary>
    ///     Builds the Excel spreadsheet from the provided list of records.
    /// </summary>
    /// <param name="records">The list of records to write to the spreadsheet.</param>
    /// <returns>The generated <see cref="IWorkbook" />.</returns>
    public IWorkbook Build(IList<TRow> records)
    {
        // Get the column names and properties from the typed object using reflection and the ExcelColumnAttribute.
        var columns = BuildColumns();
        ReorderColumns(columns);

        // Create the header row in the Excel sheet, setting cell values to the column name defined in the
        // ExcelColumnAttribute or the property name.
        var headerRow = Worksheet.CreateRow(0);
        for (var i = 0; i < ColumnDefinitions.Count; i++)
        {
            var headerCell = headerRow.CreateCell(i);
            var columnDefinition = ColumnDefinitions[i];
            headerCell.CellStyle = columnDefinition.ResolvedStyle;
            headerCell.SetCellValue(columnDefinition.ColumnName);
        }

        // Iterate through the list of records to write data rows.
        for (var i = 0; i < records.Count; i++)
        {
            // Create a new row in the Excel sheet for each record and get the current record from the list.
            var sheetRow = Worksheet.CreateRow(i + 1);
            var dataRow = records[i];

            // Iterate through the columns and write cell values for each property.
            for (var j = 0; j < ColumnDefinitions.Count; j++)
            {
                // Get the PropertyInfo for the current column's property, throwing if property doesn't exist.
                var propertyInfo = typeof(TRow).GetProperty(ColumnDefinitions[j].PropertyName);
                ArgumentNullException.ThrowIfNull(propertyInfo);

                // Get the value of the property from the current record, skipping if the value is null.
                var value = propertyInfo.GetValue(dataRow);
                if (value is null) continue;

                // Create a new cell, set the specified style if one exists, and set the cell's value.
                var cell = sheetRow.CreateCell(j);
                if (ColumnDefinitions[j].ResolvedStyle is not null) cell.CellStyle = ColumnDefinitions[j].ResolvedStyle;
                SetValue(ref cell, value);
            }
        }

        return Workbook;
    }

    /// <summary>
    ///     Gets the column names and property names from the typed object using reflection and the
    ///     <see cref="ColumnAttribute" />.
    /// </summary>
    /// <returns>A list of <see cref="ColumnInfo" /> objects containing property names, column names, and order.</returns>
    private List<ColumnInfo> BuildColumns()
    {
        // Initialize an empty list to store column definitions and iterate through each property of the generic row type.
        var columnDefinitions = new List<ColumnInfo>();
        foreach (var property in typeof(TRow).GetProperties())
        {
            var attribute = property.GetCustomAttribute<ColumnAttribute>();

            // Check if a column definition already exists for this property.
            var columnDefinition = ColumnDefinitions.FirstOrDefault(x => x.PropertyName == property.Name);
            if (columnDefinition is null)
            {
                if (attribute is null) columnDefinition = new ColumnInfo(property.Name);
                else
                    columnDefinition = new ColumnInfo(property.Name, attribute.Name, attribute.Order)
                    {
                        NumericFormat = attribute.NumericFormat
                    };
            }

            // Build the final ColumnInfo object with style and format.
            columnDefinitions.Add(BuildColumn(columnDefinition));
        }

        return columnDefinitions;
    }
}