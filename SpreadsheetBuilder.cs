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

using System.Data;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SpreadsheetHelper.Internal;

namespace SpreadsheetHelper;

/// <summary>
///     Builds an Excel spreadsheet for a <see cref="System.Data.DataTable" />.
/// </summary>
public class SpreadsheetBuilder : ImplSpreadsheetBuilder
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
    public static SpreadsheetBuilder Open(string path, string sheetName = SheetName)
    {
        // Open the Excel file for reading.
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read);
        
        // Determine the workbook type based on the file extension.
        IWorkbook workbook;
        if (path.EndsWith(".xlsx")) workbook = new XSSFWorkbook(stream);
        else if (path.EndsWith(".xls")) workbook = new HSSFWorkbook(stream);
        else throw new SpreadsheetException("Unsupported file extension. Only .xlsx and .xls are supported.");
        
        // Create a new instance of the builder, load the specified Excel file and sheet, return the builder.
        return new SpreadsheetBuilder(workbook, sheetName);
    }

    /// <summary>
    ///     Asynchronously opens an existing Excel file and initializes a <see cref="SpreadsheetBuilder{TRow}" /> for writing
    ///     to a specific sheet.
    /// </summary>
    /// <param name="path">The path to the Excel file.</param>
    /// <param name="sheetName">The name of the sheet to work with. Defaults to "Sheet1".</param>
    /// <returns>A new instance of <see cref="SpreadsheetBuilder{TRow}" />.</returns>
    public static async Task<SpreadsheetBuilder> OpenAsync(string path, string sheetName = SheetName)
    {
        // Open the Excel file for reading.
        await using var stream = new FileStream(path, FileMode.Open, FileAccess.Read);
        
        // Determine the workbook type based on the file extension.
        IWorkbook workbook;
        if (path.EndsWith(".xlsx")) workbook = new XSSFWorkbook(stream);
        else if (path.EndsWith(".xls")) workbook = new HSSFWorkbook(stream);
        else throw new SpreadsheetException("Unsupported file extension. Only .xlsx and .xls are supported.");
        
        // Create a new instance of the builder, load the specified Excel file and sheet, return the builder.
        return new SpreadsheetBuilder(workbook, sheetName);
    }

    /// <summary>
    ///     Sets the style for a specific column by its name.
    /// </summary>
    /// <param name="columnName">The name of the column to style.</param>
    /// <param name="configureStyle">An action to configure the cell style using the <see cref="CellStyleBuilder" />.</param>
    /// <returns>The current <see cref="SpreadsheetBuilder" /> instance for method chaining.</returns>
    public new SpreadsheetBuilder SetColumnStyle(string columnName, Action<CellStyleBuilder> configureStyle)
    {
        base.SetColumnStyle(columnName, configureStyle);
        return this;
    }

    /// <summary>
    ///     Sets the numeric format for a specific column by its name.
    /// </summary>
    /// <param name="columnName">The name of the column to format.</param>
    /// <param name="numericFormat">The numeric format string to apply (e.g., "0.00", "#,##0").</param>
    /// <returns>The current <see cref="SpreadsheetBuilder" /> instance for method chaining.</returns>
    public new SpreadsheetBuilder SetColumnFormat(string columnName, string numericFormat)
    {
        base.SetColumnFormat(columnName, numericFormat);
        return this;
    }

    /// <summary>
    ///     Builds the Excel spreadsheet from the provided <see cref="System.Data.DataTable" />.
    /// </summary>
    /// <param name="dataTable">The DataTable containing the data to write to the spreadsheet.</param>
    /// <returns>The generated <see cref="IWorkbook" />.</returns>
    public IWorkbook Build(DataTable dataTable)
    {
        // Get the column names from the DataTable and reorder the columns based on their original order in the DataTable.
        var columns = BuildColumns(ref dataTable);
        ReorderColumns(columns);

        // Create the header row in the Excel sheet using the DataTable's column names.
        var headerRow = Worksheet.CreateRow(0);
        for (var i = 0; i < ColumnDefinitions.Count; i++)
        {
            var headerCell = headerRow.CreateCell(i);
            var columnDefinition = ColumnDefinitions[i];
            headerCell.CellStyle = columnDefinition.ResolvedStyle;
            headerCell.SetCellValue(columnDefinition.ColumnName);
        }

        // Create data rows from the DataTable's rows.
        for (var i = 0; i < dataTable.Rows.Count; i++)
        {
            // Create a new row in the sheet and get the current DataRow from the DataTable.
            var sheetRow = Worksheet.CreateRow(i + 1);
            var dataRow = dataTable.Rows[i];

            // Iterate through the cells in the DataRow and add their values to the sheet row.
            for (var j = 0; j < dataRow.ItemArray.Length; j++)
            {
                var cell = sheetRow.CreateCell(j);
                if (ColumnDefinitions[j].ResolvedStyle is not null)
                    cell.CellStyle = ColumnDefinitions[j].ResolvedStyle;
                SetValue(ref cell, dataRow[j]);
            }
        }

        return Workbook;
    }

    /// <summary>
    ///     Gets the column names from the <see cref="System.Data.DataTable" />.
    /// </summary>
    /// <param name="dataTable">A reference to the DataTable.</param>
    /// <returns>A list of <see cref="ColumnInfo" /> objects containing column names and their original order.</returns>
    private List<ColumnInfo> BuildColumns(ref DataTable dataTable)
    {
        var columnDefinitions = new List<ColumnInfo>();
        for (var i = 0; i < dataTable.Columns.Count; i++)
        {
            // Get the column name from the DataTable's column. If column definition does not exist, create a new one.
            var columnName = dataTable.Columns[i].ColumnName;
            var columnDefinition = ColumnDefinitions.FirstOrDefault(x => x.ColumnName == columnName) ??
                                   new ColumnInfo(columnName);

            // Set the order of the column to its original index in the DataTable and build the final ColumnInfo object.
            columnDefinition.Order = i;
            columnDefinitions.Add(BuildColumn(columnDefinition));
        }

        return columnDefinitions;
    }
}