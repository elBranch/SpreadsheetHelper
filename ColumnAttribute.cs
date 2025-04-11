using System.Diagnostics;

namespace ExcelHelper;

/// <summary>
///     Attribute used to specify the column name and order in an Excel sheet.
/// </summary>
[DebuggerDisplay("{Name} Order: {Order} Format: {NumericFormat}")]
[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
public sealed class ColumnAttribute : Attribute
{
    private int _order = -1;

    /// <summary>
    ///     Initializes a new ExcelColumnAttribute.
    /// </summary>
    public ColumnAttribute()
    {
    }

    /// <summary>
    ///     Initializes a new ExcelColumnAttribute with the specified column name.
    /// </summary>
    /// <param name="name">The name of the column in the Excel sheet.</param>
    public ColumnAttribute(string name)
    {
        Name = name;
    }

    /// <summary>Sets the name of the column in the Excel sheet.</summary>
    public string? Name { get; private set; }

    /// <summary>Gets or sets the format of a numeric cell.</summary>
    public string? NumericFormat { get; set; }

    /// <summary>Gets or sets the order of the column in the Excel sheet.</summary>
    public int Order
    {
        get => _order;
        set
        {
            if (value < 0) throw new SpreadsheetException("Order must be greater than zero.");
            _order = value;
        }
    }
}