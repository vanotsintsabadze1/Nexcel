using ClosedXML.Excel;
using Nexcel.Misc;
using Nexcel.Models;

namespace Nexcel.Interfaces;

public interface IInitializedWorksheetStage : IDisposable
{
    IInitializedWorksheetStage InitializeByCollection<T>(
    IEnumerable<T> items,
    string? headerInitializationStartCell = ExcelBuilderConstants.DefaultHeadersInitializationCell,
    string? initializationStartCell = ExcelBuilderConstants.DefaultItemInitializationCell);

    IInitializedWorksheetStage SetFontSize(string cell, int fontSize);

    IInitializedWorksheetStage SetGlobalFontSize(int fontSize);

    IInitializedWorksheetStage SetFontName(string fontName);

    IInitializedWorksheetStage SetBackgroundColor(string cell, XLColor backgroundColor);

    IInitializedWorksheetStage SetBorderStyle(string cell, XLBorder border, XLBorderStyleValues borderStyleValue);

    IInitializedWorkbookStage Done();
}
