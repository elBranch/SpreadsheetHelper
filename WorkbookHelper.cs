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

using NPOI.SS.UserModel;

namespace SpreadsheetHelper;

/// <summary>
///     Provides extension methods for saving Excel workbooks.
/// </summary>
public static class WorkbookHelper
{
    /// <summary>
    ///     Saves the Excel workbook to the specified file path synchronously.
    /// </summary>
    /// <param name="workbook">The Excel workbook to save.</param>
    /// <param name="path">The file path to save the workbook to.</param>
    public static void SaveAs(this IWorkbook workbook, string path)
    {
        using var stream = new FileStream(path, FileMode.Create, FileAccess.Write);
        workbook.Write(stream);
    }

    /// <summary>
    ///     Saves the Excel workbook to the specified file path asynchronously.
    /// </summary>
    /// <param name="workbook">The Excel workbook to save.</param>
    /// <param name="path">The file path to save the workbook to.</param>
    public static async Task SaveAsAsync(this IWorkbook workbook, string path)
    {
        await using var stream = new FileStream(path, FileMode.Create, FileAccess.Write);
        workbook.Write(stream);
    }
}