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

namespace SpreadsheetHelper;

/// <summary>
///     Attribute used to specify the column name and order in an Excel sheet.
/// </summary>
[DebuggerDisplay("{Name} Order: {Order} Format: {NumericFormat}")]
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public sealed class ColumnAttribute : Attribute
{
    /// <summary>
    ///     Initializes a new ExcelColumnAttribute.
    /// </summary>
    public ColumnAttribute()
    {
    }

    /// <summary>
    ///     Initializes a new ExcelColumnAttribute with the specified column name.
    /// </summary>
    /// <param name="name">The column header (or name) as displayed in the spreadsheet.</param>
    public ColumnAttribute(string name)
    {
        Name = name;
    }

    /// <summary>Sets the name of the column in the Excel sheet.</summary>
    public string? Name { get; }

    /// <summary>Gets or sets the format of a numeric cell.</summary>
    public string? NumericFormat { get; set; }

    /// <summary>Gets or sets the order of the column in the Excel sheet.</summary>
    public int Order { get; set; }
}