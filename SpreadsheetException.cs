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

namespace SpreadsheetHelper;

/// <summary>
///     Exception type used for errors related to Excel operations.
/// </summary>
public class SpreadsheetException : ApplicationException
{
    /// <summary>
    ///     Initializes a new instance of the <see cref="SpreadsheetException" /> class with a specified error message.
    /// </summary>
    /// <param name="message">The error message that describes the exception.</param>
    public SpreadsheetException(string message) : base(message)
    {
    }

    /// <summary>
    ///     Initializes a new instance of the <see cref="SpreadsheetException" /> class with a specified error message and a
    ///     reference to the inner exception that is the cause of this exception.
    /// </summary>
    /// <param name="message">The error message that describes the exception.</param>
    /// <param name="innerException">The exception that is the cause of the current exception.</param>
    public SpreadsheetException(string message, Exception innerException) : base(message, innerException)
    {
    }
}