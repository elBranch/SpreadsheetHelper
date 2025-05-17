using NPOI.SS.UserModel;

namespace SpreadsheetHelper;

/// <summary>
///     Provides extension methods for converting <see cref="IWorkbook" /> and <see cref="ISheet" /> instances to
///     <see cref="Spreadsheet{TEntity}" /> helpers for strongly-typed spreadsheet operations.
/// </summary>
public static class WorkbookHelpers
{
    /// <summary>
    ///     Converts the specified <see cref="IWorkbook" /> to a <see cref="Spreadsheet{TEntity}" /> helper, using the provided
    ///     sheet name or the default sheet name if not specified.
    /// </summary>
    /// <typeparam name="TEntity">
    ///     The type of entity to map spreadsheet rows to. Must be a class with a parameterless constructor.
    /// </typeparam>
    /// <param name="workbook">The workbook to convert.</param>
    /// <param name="sheetName">The name of the sheet to use. Defaults to <see cref="Spreadsheet{TEntity}.DefaultSheetName" />.</param>
    /// <returns>A <see cref="Spreadsheet{TEntity}" /> instance for the specified workbook and sheet.</returns>
    public static Spreadsheet<TEntity> ToHelper<TEntity>(this IWorkbook workbook,
        string sheetName = Spreadsheet<TEntity>.DefaultSheetName) where TEntity : class, new()
    {
        return new Spreadsheet<TEntity>(workbook, sheetName);
    }

    /// <summary>
    ///     Converts the specified <see cref="ISheet" /> to a <see cref="Spreadsheet{TEntity}" /> helper, using the sheet's
    ///     workbook and name.
    /// </summary>
    /// <typeparam name="TEntity">
    ///     The type of entity to map spreadsheet rows to. Must be a class with a parameterless constructor.
    /// </typeparam>
    /// <param name="sheet">The sheet to convert.</param>
    /// <returns>A <see cref="Spreadsheet{TEntity}" /> instance for the specified sheet.</returns>
    public static Spreadsheet<TEntity> ToHelper<TEntity>(this ISheet sheet) where TEntity : class, new()
    {
        return new Spreadsheet<TEntity>(sheet.Workbook, sheet.SheetName);
    }
    
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