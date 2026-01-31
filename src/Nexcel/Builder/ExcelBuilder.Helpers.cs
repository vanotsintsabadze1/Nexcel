using ClosedXML.Excel;

namespace Nexcel.Builder;

public partial class ExcelBuilder
{
    private void ThrowIfWorksheetDoesNotExist()
    {
        if (this._focusedWorksheet == null)
            throw new InvalidOperationException("Worksheet was not found");
    }

    private IXLCell GetCell(string cell)
    {
        var existingCell = this._focusedWorksheet.Cell(cell);

        if (existingCell == null)
            throw new InvalidOperationException($"Can not perform operation {nameof(SetFontSize)} on non-existing cell {cell}");

        return existingCell;
    }

    private void ThrowIfWorkbookIsDisposed()
    {
        if (_workbook == null)
            throw new InvalidOperationException("Workbook instance has been disposed already");
    }

    private IXLWorksheet GetSheetByName(string sheetName)
    {
        var sheet = this._workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);

        if (sheet == null)
            throw new ArgumentException($"Sheet with name {sheetName} does not exist");

        return sheet;
    }

    private IXLWorksheet GetSheetByIndex(int sheetIndex)
    {
        try
        {
            return this._workbook.Worksheets.ToList()[sheetIndex];
        }
        catch (ArgumentOutOfRangeException ex)
        {
            throw new ArgumentException($"No sheet was found with index of: {sheetIndex}");
        }
    }
}
