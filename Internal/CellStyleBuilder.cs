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
using HorizontalAlignment = NPOI.SS.UserModel.HorizontalAlignment;
using VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment;

namespace SpreadsheetHelper.Internal;

/// <summary>
///     Provides a fluent interface for building and configuring cell styles.
/// </summary>
public class CellStyleBuilder
{
    private readonly IWorkbook _workbook;

    /// <summary>
    ///     The underlying cell style.
    /// </summary>
    internal readonly ICellStyle CellStyle;

    /// <summary>
    ///     Initializes a new instance of the <see cref="CellStyleBuilder" /> class.
    /// </summary>
    /// <param name="workbook">The Excel workbook.</param>
    /// <param name="cellStyle">The cell style to build.</param>
    public CellStyleBuilder(IWorkbook workbook, ICellStyle cellStyle)
    {
        _workbook = workbook;
        CellStyle = cellStyle;
    }

    /// <summary>
    ///     Sets the horizontal alignment of the cell.
    /// </summary>
    /// <param name="horizontalAlignment">The horizontal alignment to set.</param>
    /// <returns>The current <see cref="CellStyleBuilder" /> instance for method chaining.</returns>
    public CellStyleBuilder Alignment(HorizontalAlignment horizontalAlignment)
    {
        return Alignment(horizontalAlignment, VerticalAlignment.None);
    }

    /// <summary>
    ///     Sets the vertical alignment of the cell.
    /// </summary>
    /// <param name="verticalAlignment">The vertical alignment to set.</param>
    /// <returns>The current <see cref="CellStyleBuilder" /> instance for method chaining.</returns>
    public CellStyleBuilder Alignment(VerticalAlignment verticalAlignment)
    {
        return Alignment(HorizontalAlignment.General, verticalAlignment);
    }

    /// <summary>
    ///     Sets both the horizontal and vertical alignment of the cell.
    /// </summary>
    /// <param name="horizontalAlignment">The horizontal alignment to set.</param>
    /// <param name="verticalAlignment">The vertical alignment to set.</param>
    /// <returns>The current <see cref="CellStyleBuilder" /> instance for method chaining.</returns>
    public CellStyleBuilder Alignment(HorizontalAlignment horizontalAlignment, VerticalAlignment verticalAlignment)
    {
        CellStyle.Alignment = horizontalAlignment;
        CellStyle.VerticalAlignment = verticalAlignment;
        return this;
    }

    /// <summary>
    ///     Sets the border color for all borders of the cell.
    /// </summary>
    /// <param name="borderColor">The color to set for the borders.</param>
    /// <returns>The current <see cref="CellStyleBuilder" /> instance for method chaining.</returns>
    public CellStyleBuilder BorderColor(IColor borderColor)
    {
        return BorderColor(borderColor, borderColor);
    }

    /// <summary>
    ///     Sets the border color for the top/bottom and right/left borders of the cell.
    /// </summary>
    /// <param name="topBottomBorder">The color to set for the top and bottom borders.</param>
    /// <param name="rightLeftBorder">The color to set for the right and left borders.</param>
    /// <returns>The current <see cref="CellStyleBuilder" /> instance for method chaining.</returns>
    public CellStyleBuilder BorderColor(IColor topBottomBorder, IColor rightLeftBorder)
    {
        return BorderColor(topBottomBorder, rightLeftBorder, topBottomBorder);
    }

    /// <summary>
    ///     Sets the border color for the top, right/left, and bottom borders of the cell.
    /// </summary>
    /// <param name="topBorder">The color to set for the top border.</param>
    /// <param name="rightLeftBorder">The color to set for the right and left borders.</param>
    /// <param name="bottomBorder">The color to set for the bottom border.</param>
    /// <returns>The current <see cref="CellStyleBuilder" /> instance for method chaining.</returns>
    public CellStyleBuilder BorderColor(IColor topBorder, IColor rightLeftBorder, IColor bottomBorder)
    {
        return BorderColor(topBorder, rightLeftBorder, bottomBorder, rightLeftBorder);
    }

    /// <summary>
    ///     Sets the border color for each border of the cell individually.
    /// </summary>
    /// <param name="topBorder">The color to set for the top border.</param>
    /// <param name="rightBorder">The color to set for the right border.</param>
    /// <param name="bottomBorder">The color to set for the bottom border.</param>
    /// <param name="leftBorder">The color to set for the left border.</param>
    /// <returns>The current <see cref="CellStyleBuilder" /> instance for method chaining.</returns>
    public CellStyleBuilder BorderColor(IColor topBorder, IColor rightBorder, IColor bottomBorder,
        IColor leftBorder)
    {
        CellStyle.TopBorderColor = topBorder.Indexed;
        CellStyle.RightBorderColor = rightBorder.Indexed;
        CellStyle.BottomBorderColor = bottomBorder.Indexed;
        CellStyle.LeftBorderColor = leftBorder.Indexed;
        return this;
    }

    /// <summary>
    ///     Sets the border style for all borders of the cell.
    /// </summary>
    /// <param name="borderStyle">The border style to set.</param>
    /// <returns>The current <see cref="CellStyleBuilder" /> instance for method chaining.</returns>
    public CellStyleBuilder BorderStyle(BorderStyle borderStyle)
    {
        return BorderStyle(borderStyle, borderStyle);
    }

    /// <summary>
    ///     Sets the border style for the top/bottom and right/left borders of the cell.
    /// </summary>
    /// <param name="topBottomBorder">The border style to set for the top and bottom borders.</param>
    /// <param name="rightLeftBorder">The border style to set for the right and left borders.</param>
    /// <returns>The current <see cref="CellStyleBuilder" /> instance for method chaining.</returns>
    public CellStyleBuilder BorderStyle(BorderStyle topBottomBorder, BorderStyle rightLeftBorder)
    {
        return BorderStyle(topBottomBorder, rightLeftBorder, topBottomBorder);
    }

    /// <summary>
    ///     Sets the border style for the top, right/left, and bottom borders of the cell.
    /// </summary>
    /// <param name="topBorder">The border style to set for the top border.</param>
    /// <param name="rightLeftBorder">The border style to set for the right and left borders.</param>
    /// <param name="bottomBorder">The border style to set for the bottom border.</param>
    /// <returns>The current <see cref="CellStyleBuilder" /> instance for method chaining.</returns>
    public CellStyleBuilder BorderStyle(BorderStyle topBorder, BorderStyle rightLeftBorder, BorderStyle bottomBorder)
    {
        return BorderStyle(topBorder, rightLeftBorder, bottomBorder, rightLeftBorder);
    }

    /// <summary>
    ///     Sets the border style for each border of the cell individually.
    /// </summary>
    /// <param name="topBorder">The border style to set for the top border.</param>
    /// <param name="rightBorder">The border style to set for the right border.</param>
    /// <param name="bottomBorder">The border style to set for the bottom border.</param>
    /// <param name="leftBorder">The border style to set for the left border.</param>
    /// <returns>The current <see cref="CellStyleBuilder" /> instance for method chaining.</returns>
    public CellStyleBuilder BorderStyle(BorderStyle topBorder, BorderStyle rightBorder, BorderStyle bottomBorder,
        BorderStyle leftBorder)
    {
        CellStyle.BorderTop = topBorder;
        CellStyle.BorderRight = rightBorder;
        CellStyle.BorderBottom = bottomBorder;
        CellStyle.BorderLeft = leftBorder;
        return this;
    }

    /// <summary>
    ///     Sets the background color and fill pattern of the cell.
    /// </summary>
    /// <param name="color">The background color to set.</param>
    /// <param name="fillPattern">The fill pattern to use. Defaults to <see cref="FillPattern.SolidForeground" />.</param>
    /// <returns>The current <see cref="CellStyleBuilder" /> instance for method chaining.</returns>
    public CellStyleBuilder BackgroundColor(IColor color, FillPattern fillPattern = FillPattern.SolidForeground)
    {
        CellStyle.FillBackgroundColor = color.Indexed;
        CellStyle.FillPattern = fillPattern;
        return this;
    }

    /// <summary>
    ///     Sets the foreground color and fill pattern of the cell.
    /// </summary>
    /// <param name="color">The foreground color to set.</param>
    /// <param name="fillPattern">The fill pattern to use.</param>
    /// <returns>The current <see cref="CellStyleBuilder" /> instance for method chaining.</returns>
    public CellStyleBuilder ForegroundColor(IColor color, FillPattern fillPattern)
    {
        CellStyle.FillForegroundColor = color.Indexed;
        CellStyle.FillPattern = fillPattern;
        return this;
    }

    /// <summary>
    ///     Configures the font of the cell.
    /// </summary>
    /// <param name="fontName">The name of the font. If null, the existing font name is used.</param>
    /// <param name="fontSize">The size of the font. If null, the existing font size is used.</param>
    /// <param name="color">The color of the font. If null, the existing font color is used.</param>
    /// <param name="bold">Whether the font should be bold. If null, the existing bold setting is used.</param>
    /// <param name="italic">Whether the font should be italic. If null, the existing italic setting is used.</param>
    /// <param name="strikeout">Whether the font should be strikeout. If null, the existing strikeout setting is used.</param>
    /// <param name="underline">The underline type of the font. If null, the existing underline type is used.</param>
    /// <param name="superScript">The super script type of the font. If null, the existing super script type is used.</param>
    /// <returns>The current <see cref="CellStyleBuilder" /> instance for method chaining.</returns>
    public CellStyleBuilder Font(string? fontName = null, short? fontSize = null, IColor? color = null,
        bool? bold = null, bool? italic = null, bool? strikeout = null, FontUnderlineType? underline = null,
        FontSuperScript? superScript = null)
    {
        var baseFont = CellStyle.GetFont(_workbook);

        var qIsBold = bold ?? baseFont.IsBold;
        var qColor = color?.Indexed ?? baseFont.Color;
        var qFontHeight = fontSize ?? Convert.ToInt16(baseFont.FontHeight);
        var qFontName = fontName ?? baseFont.FontName;
        var qIsItalic = italic ?? baseFont.IsItalic;
        var qIsStrikeout = strikeout ?? baseFont.IsStrikeout;
        var qTypeOffset = superScript ?? baseFont.TypeOffset;
        var qUnderline = underline ?? baseFont.Underline;

        var font = _workbook.FindFont(qIsBold, qColor, qFontHeight, qFontName, qIsItalic, qIsStrikeout, qTypeOffset,
            qUnderline) ?? _workbook.CreateFont();
        font.IsBold = qIsBold;
        font.Color = qColor;
        font.FontHeight = qFontHeight;
        font.FontName = qFontName;
        font.IsItalic = qIsItalic;
        font.IsStrikeout = qIsStrikeout;
        font.TypeOffset = qTypeOffset;
        font.Underline = qUnderline;

        CellStyle.SetFont(font);
        return this;
    }
}