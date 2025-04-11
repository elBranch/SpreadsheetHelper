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

using System.Diagnostics;
using NPOI.SS.UserModel;

namespace ExcelHelper.Internal;

/// <summary>
///     Represents a column in the Excel sheet, containing property name, column name, and order.
/// </summary>
[DebuggerDisplay("{ColumnName}")]
internal class ColumnInfo
{
    /// <summary>
    ///     Initializes a new ExcelColumn with the property name as the column name and default order.
    /// </summary>
    /// <param name="propertyName">The name of the property.</param>
    public ColumnInfo(string propertyName)
    {
        PropertyName = propertyName;
        ColumnName = propertyName;
        Order = -1;
    }

    /// <summary>
    ///     Initializes a new ExcelColumn with the specified property name, column name, and order.
    /// </summary>
    /// <param name="propertyName">The name of the property.</param>
    /// <param name="columnName">The name of the column in the Excel sheet.</param>
    /// <param name="order">The order of the column in the Excel sheet.</param>
    public ColumnInfo(string propertyName, string? columnName, int order = -1)
    {
        PropertyName = propertyName;
        ColumnName = columnName ?? propertyName;
        Order = order;
    }

    /// <summary>Gets the name of the property.</summary>
    public string PropertyName { get; }

    /// <summary>Gets the name of the column in the Excel sheet.</summary>
    public string ColumnName { get; }

    /// <summary>Gets the order of the column in the Excel sheet.</summary>
    public int Order { get; internal set; }

    /// <summary>
    ///     Defines the visual formatting and style properties of an Excel cell, including font, alignment, borders, fill,
    ///     and data format.
    /// </summary>
    public string? NumericFormat { get; internal set; }

    /// <summary>
    ///     An action to configure the cell style.
    /// </summary>
    internal Action<ICellStyle>? StyleConfiguration { get; set; }

    /// <summary>
    ///     The resolved cell style.
    /// </summary>
    internal ICellStyle? ResolvedStyle { get; set; }
}