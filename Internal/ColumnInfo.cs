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
using System.Reflection;
using NPOI.SS.UserModel;

namespace SpreadsheetHelper.Internal;

/// <summary>
///     Represents a column in the Excel sheet, containing property name, column name, and order.
/// </summary>
[DebuggerDisplay("{ColumnName}")]
internal struct ColumnInfo
{
    /// <summary>
    ///     Gets the name of the column in the Excel sheet.
    /// </summary>
    public required string ColumnName { get; init; }
    
    /// <summary>
    ///     Identifies the type from the mapped object.
    /// </summary>
    internal PropertyInfo Property { get; init; }

    /// <summary>
    ///     Gets the order of the column in the Excel sheet.
    /// </summary>
    public required int Order { get; init; }

    /// <summary>
    ///     Defines the visual formatting and style properties of an Excel cell, including font, alignment, borders, fill, and
    ///     data format.
    /// </summary>
    // TODO: We need to use just this and get rid of StyleConfiguration and ResolvedStyle. If we want to build in styles then work on that in the next major version. Implement it in the Spreadsheet<TRow>.Rows and/or Spreadsheet<TRow>.SetRecord
    public string? NumericFormat { get; init; }

    /// <summary>
    ///     An action to configure the cell style.
    /// </summary>
    internal Action<ICellStyle>? StyleConfiguration { get; set; }

    /// <summary>
    ///     The resolved cell style.
    /// </summary>
    internal ICellStyle? ResolvedStyle { get; set; }
}