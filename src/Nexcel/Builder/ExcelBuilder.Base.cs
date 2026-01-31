using ClosedXML.Excel;
using Nexcel.Interfaces;
using Nexcel.Misc;
using Nexcel.Models;
using Nexcel.Utilities;

namespace Nexcel.Builder;

public partial class ExcelBuilder : IInitializedWorkbookStage, IInitializedWorksheetStage
{
    private IXLWorkbook _workbook;
    private IXLWorksheet _focusedWorksheet;

    internal ExcelBuilder(IXLWorkbook workbook) => this._workbook = workbook;
    
    public IInitializedWorksheetStage AddWorksheet(string? sheetName = "Sheet1")
    {
        this._focusedWorksheet = this._workbook.AddWorksheet(sheetName);
        return this;
    }

    public IInitializedWorksheetStage InitializeByCollection<T>(
        IEnumerable<T> items, 
        string? headerInitializationStartCell = ExcelBuilderConstants.DefaultHeadersInitializationCell,
        string? initializationStartCell = ExcelBuilderConstants.DefaultItemInitializationCell)
    {
        var itemType = items.FirstOrDefault()!.GetType();
        var propertyNames = ReflectionUtility.ExtractPropertyNames(itemType);

        ItemInitializationUtility.InitializeWorksheetHeaders(this._focusedWorksheet, propertyNames);
        ItemInitializationUtility.InitializeWorksheetItems(this._focusedWorksheet, items);

        return this;
    }

    public IInitializedWorkbookStage RemoveWorksheet(string? sheetName = null)
    {
        if (sheetName == null)
        {
            switch (_focusedWorksheet)
            {
                case null:
                    throw new InvalidOperationException("No current focused worksheet exists. \n This is due to removing a worksheet previously and not focusing on other sheet or there was no worksheet at all");
                case IXLWorksheet:
                    this._workbook.Worksheets.Delete(_focusedWorksheet.Name);
                    return this;
                default:
                    throw new InvalidOperationException($"Something went wrong while trying to remove a worksheet {sheetName}. If sheet name is null, then the active worksheet was not found");
            }
        }

        var sheet = this._workbook.Worksheets.FirstOrDefault(ws => ws.Name == sheetName);

        if (sheet != null)
        {
            this._workbook.Worksheets.Delete(sheetName);
            return this;
        }

        throw new ArgumentException($"No worksheet was found with name {sheetName}");
    }
    
    public IInitializedWorksheetStage SelectWorksheet(string sheetName)
    {
        this._focusedWorksheet = GetSheetByName(sheetName);
        return this;
    }

    public IInitializedWorksheetStage SelectWorksheet(int sheetIndex)
    {
        this._focusedWorksheet = GetSheetByIndex(sheetIndex);
        return this;
    }

    public IInitializedWorkbookStage Done() => this;

    public byte[] Build()
    {
        ThrowIfWorkbookIsDisposed();

        using var memoryStream = new MemoryStream();

        memoryStream.Seek(0, SeekOrigin.Begin);
        
        this._workbook.SaveAs(memoryStream);
        this.Dispose();

        return memoryStream.ToArray();
    }

    public void BuildAndSave(string path)
    {
        ThrowIfWorkbookIsDisposed();

        this._workbook.SaveAs(path);
        this.Dispose();
    }

    public void Dispose() => this._workbook.Dispose();
}