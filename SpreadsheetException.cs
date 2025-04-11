namespace ExcelHelper;

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