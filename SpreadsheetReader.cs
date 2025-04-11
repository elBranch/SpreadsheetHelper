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
using System.Reflection;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SpreadsheetHelper.Internal;

namespace SpreadsheetHelper;

/// <summary>
///     Provides methods for reading data from Excel files.
/// </summary>
public class SpreadsheetReader
{
    /// <summary>
    ///     Asynchronously reads data from an Excel file and returns it as a <see cref="System.Data.DataTable" />.
    /// </summary>
    /// <param name="filePath">The path to the Excel file.</param>
    /// <param name="sheetName">The name of the sheet to read from. If null, the active sheet is used.</param>
    /// <returns>A <see cref="System.Data.DataTable" /> containing the data from the Excel file.</returns>
    /// <exception cref="SpreadsheetException">
    ///     Thrown if the file extension is unsupported, the specified sheet does not exist, or the header row is missing.
    /// </exception>
    public static async Task<DataTable> ReadFileAsync(string filePath,
        string? sheetName = ImplSpreadsheetBuilder.SheetName)
    {
        var dataTable = new DataTable();

        // Open the Excel file for reading.
        await using var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read);

        // Determine the workbook type based on the file extension.
        IWorkbook workbook;
        if (filePath.EndsWith(".xlsx")) workbook = new XSSFWorkbook(stream);
        else if (filePath.EndsWith(".xls")) workbook = new HSSFWorkbook(stream);
        else throw new SpreadsheetException("Unsupported file extension. Only .xlsx and .xls are supported.");

        // Get the specified sheet or the active sheet if no sheet name is provided.
        var sheet = sheetName is null ? workbook.GetSheet(sheetName) : workbook.GetSheetAt(workbook.ActiveSheetIndex);
        if (sheet is null) throw new SpreadsheetException($"Sheet name '{sheetName}' does not exist.");

        // Get the header row from the sheet.
        var headerRow = sheet.GetRow(sheet.FirstRowNum);
        if (headerRow == null) throw new SpreadsheetException("Header row is missing.");

        // Create header from the first row of the sheet, using the cell values as columns in the DataTable.
        var firstRow = sheet.GetRow(sheet.FirstRowNum + 1);
        for (var i = 0; i < headerRow.LastCellNum; i++)
        {
            var headerCell = headerRow.GetCell(i);
            if (headerCell == null || string.IsNullOrEmpty(headerCell.StringCellValue)) continue;

            var columnName = headerCell.StringCellValue.Trim();
            dataTable.Columns.Add(columnName, GetColumnType(firstRow.Cells[i]));
        }

        // Create rows from the remaining rows of the sheet.
        for (var rowIndex = sheet.FirstRowNum + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            // Get the current row from the sheet and create a DataRow for the DataTable.
            var sheetRow = sheet.GetRow(rowIndex);
            if (sheetRow == null) continue;

            var record = dataTable.NewRow();
            for (var cellIndex = sheetRow.FirstCellNum; cellIndex < sheetRow.LastCellNum; cellIndex++)
            {
                var cell = sheetRow.GetCell(cellIndex);
                if (cell is null) continue;

                var cellValue = GetCellValue(cell);
                if (cellValue is null) continue;
                
                record[cellIndex] = cellValue;
            }

            dataTable.Rows.Add(record);
        }

        return dataTable;
    }

    /// <summary>
    ///     Asynchronously reads data from an Excel file and maps it to a list of objects of type <typeparamref name="TRow" />.
    /// </summary>
    /// <typeparam name="TRow">The type to map the Excel data to. Must have a parameterless constructor.</typeparam>
    /// <param name="filePath">The path to the Excel file.</param>
    /// <param name="sheetName">The name of the sheet to read from. If null, the active sheet is used.</param>
    /// <returns>A list of objects of type <typeparamref name="TRow" /> containing the data from the Excel file.</returns>
    /// <exception cref="SpreadsheetException">
    ///     Thrown if the file extension is unsupported, the specified sheet does not exist, the header row is missing,
    ///     or if there is an error during data conversion or casting.
    /// </exception>
    public static async Task<List<TRow>> ReadFileAsync<TRow>(string filePath,
        string? sheetName = ImplSpreadsheetBuilder.SheetName) where TRow : new()
    {
        var records = new List<TRow>();

        // Open the Excel file for reading.
        await using var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);

        // Determine the workbook type based on the file extension.
        IWorkbook book;
        if (filePath.EndsWith(".xlsx")) book = new XSSFWorkbook(fs);
        else if (filePath.EndsWith(".xls")) book = new HSSFWorkbook(fs);
        else throw new SpreadsheetException("Unsupported file extension. Only .xlsx and .xls are supported.");

        // Get the specified sheet or the active sheet if no sheet name is provided.
        var sheet = sheetName is null ? book.GetSheet(sheetName) : book.GetSheetAt(book.ActiveSheetIndex);
        if (sheet is null) throw new SpreadsheetException($"Sheet name '{sheetName}' does not exist.");

        // Get the header row from the sheet.
        var headerRow = sheet.GetRow(sheet.FirstRowNum);
        if (headerRow == null) throw new SpreadsheetException("Header row is missing.");

        // Create a mapping of column index to property name
        var propertyMap = new Dictionary<int, PropertyInfo>();
        var properties = typeof(TRow).GetProperties();

        // Read header row and match column names to property names (case-insensitive)
        for (var i = 0; i < headerRow.LastCellNum; i++)
        {
            var headerCell = headerRow.GetCell(i);
            if (headerCell == null || string.IsNullOrEmpty(headerCell.StringCellValue)) continue;

            var columnName = headerCell.StringCellValue.Trim();
            var matchingProperty = properties.FirstOrDefault(p => IsMatchColumnName(p, columnName));

            if (matchingProperty != null) propertyMap[i] = matchingProperty;
        }

        // Read data rows (starting from the second row)
        for (var rowIndex = sheet.FirstRowNum + 1; rowIndex <= sheet.LastRowNum; rowIndex++)
        {
            var sheetRow = sheet.GetRow(rowIndex);
            if (sheetRow == null) continue;

            var record = new TRow();
            for (var cellIndex = sheetRow.FirstCellNum; cellIndex < sheetRow.LastCellNum; cellIndex++)
                if (propertyMap.TryGetValue(cellIndex, out var property))
                {
                    var cell = sheetRow.GetCell(cellIndex);
                    if (cell == null) continue;

                    var cellValue = GetCellValue(cell);
                    if (cellValue == null) continue;

                    try
                    {
                        // Determine the Type on the TRow property and attempt to apply it's value.
                        var propertyType = property.PropertyType;
                        if (propertyType.IsGenericType && propertyType.GetGenericTypeDefinition() == typeof(Nullable<>))
                            propertyType = Nullable.GetUnderlyingType(propertyType)!;
                        property.SetValue(record, Convert.ChangeType(cellValue, propertyType));
                    }
                    catch (FormatException ex)
                    {
                        throw new SpreadsheetException($"Could not convert value '{cellValue}' in column " +
                                                       $"'{GetPropertyColumnName(property)}'.", ex);
                    }
                    catch (InvalidCastException ex)
                    {
                        throw new SpreadsheetException($"Invalid cast of '{cellValue}' in column " +
                                                       $"'{GetPropertyColumnName(property)}' and row {rowIndex} to " +
                                                       $"type '{property.PropertyType.Name}'.", ex);
                    }
                }

            records.Add(record);
        }

        return records;
    }

    /// <summary>
    ///     Determines the corresponding C# type based on the provided Excel cell type.
    /// </summary>
    /// <param name="cell">The Excel cell to map to a C# type.</param>
    /// <returns>The corresponding C# type.</returns>
    private static Type GetColumnType(ICell cell)
    {
        var type = cell.CellType;
        return type switch
        {
            CellType.Numeric when DateUtil.IsCellDateFormatted(cell) => typeof(DateTime),
            CellType.Numeric => typeof(double),
            CellType.Boolean => typeof(bool),
            CellType.String => typeof(string),
            CellType.Formula => TryGetNumericValue(cell),
            CellType.Blank => typeof(string),
            CellType.Error => typeof(string),
            CellType.Unknown => typeof(string),
            _ => typeof(string)
        };
        
        // Helper function to try getting the numeric value from a formula cell.
        Type TryGetNumericValue(ICell numericCell)
        {
            try
            {
                return typeof(double);
            }
            catch (Exception)
            {
                return typeof(string);
            }
        }
    }

    /// <summary>
    ///     Gets the column name for a given property, either from the <see cref="ColumnAttribute" /> or the property name
    ///     itself.
    /// </summary>
    /// <param name="property">The <see cref="PropertyInfo" /> of the property.</param>
    /// <returns>The column name.</returns>
    private static string GetPropertyColumnName(PropertyInfo property)
    {
        var attribute = property.GetCustomAttribute<ColumnAttribute>();
        return attribute?.Name ?? property.Name;
    }

    /// <summary>
    ///     Checks if the given column name matches the property name (case-insensitive).
    /// </summary>
    /// <param name="property">The <see cref="PropertyInfo" /> of the property.</param>
    /// <param name="columnName">The column name to compare with.</param>
    /// <returns>True if the column name matches the property name, false otherwise.</returns>
    private static bool IsMatchColumnName(PropertyInfo property, string columnName)
    {
        return columnName.Equals(GetPropertyColumnName(property), StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    ///     Gets the value from an Excel cell, handling different cell types.
    /// </summary>
    /// <param name="cell">The Excel cell to get the value from.</param>
    /// <returns>The value of the cell, or null if the cell is null or blank.</returns>
    private static object? GetCellValue(ICell? cell)
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
            catch (Exception)
            {
                return numericCell.StringCellValue;
            }
        }
    }
}