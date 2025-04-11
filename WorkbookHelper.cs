using NPOI.SS.UserModel;

namespace ExcelHelper;

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