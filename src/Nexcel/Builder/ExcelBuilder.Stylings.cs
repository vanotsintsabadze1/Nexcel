using ClosedXML.Excel;
using Nexcel.Interfaces;
using Nexcel.Models;

namespace Nexcel.Builder;

public partial class ExcelBuilder
{
    public IInitializedWorksheetStage SetFontSize(string cell, int fontSize)
    {
        ThrowIfWorksheetDoesNotExist();

        var existingCell = GetCell(cell);

        existingCell.Style.Font.FontSize = fontSize;

        return this;
    }

    public IInitializedWorksheetStage SetGlobalFontSize(int fontSize)
    {
        ThrowIfWorksheetDoesNotExist();

        this._focusedWorksheet.Style.Font.FontSize = fontSize;

        return this;
    }

    public IInitializedWorksheetStage SetFontName(string fontName)
    {
        ThrowIfWorksheetDoesNotExist();

        this._focusedWorksheet.Style.Font.FontName = fontName;

        return this;
    }

    public IInitializedWorksheetStage SetBackgroundColor(string cell, XLColor backgroundColor)
    {
        ThrowIfWorksheetDoesNotExist();

        var existingCell = GetCell(cell);

        existingCell.Style.Fill.BackgroundColor = backgroundColor;
        
        return this;
    }

    public IInitializedWorksheetStage SetBorderStyle(string cell, XLBorder border, XLBorderStyleValues borderStyleValue)
    {
        var existingCell = GetCell(cell);

        switch (border)
        {
            case XLBorder.Top:
                existingCell.Style.Border.TopBorder = borderStyleValue;
                break;
            case XLBorder.Bottom:
                existingCell.Style.Border.BottomBorder = borderStyleValue;
                break;
            case XLBorder.Left:
                existingCell.Style.Border.LeftBorder = borderStyleValue;
                break;
            case XLBorder.Right:
                existingCell.Style.Border.RightBorder = borderStyleValue;
                break;
            case XLBorder.All:
                existingCell.Style.Border.SetTopBorder(borderStyleValue)
                                         .Border.SetBottomBorder(borderStyleValue)
                                         .Border.SetLeftBorder(borderStyleValue)
                                         .Border.SetRightBorder(borderStyleValue);
            break;
            default: goto case XLBorder.All;
        }

        return this;
    }
}
