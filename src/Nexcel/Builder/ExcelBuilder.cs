using ClosedXML.Excel;
using Nexcel.Helpers;
using Nexcel.Interfaces;
using Nexcel.Misc;
using Nexcel.Utilities;

namespace Nexcel.Builder;

public class ExcelBuilder : 
    IInitializedWorkbookStage, 
    IItemInitializationStage,
    IStyleableStage,
    IBuildStage,
    IDisposable
{
    private IXLWorkbook _workbook;
    private IXLWorksheet _focusedWorksheet;

    internal ExcelBuilder(IXLWorkbook workbook) => this._workbook = workbook;
    
    public IItemInitializationStage AddWorksheet(string? sheetName = "Sheet1")
    {
        this._focusedWorksheet = this._workbook.AddWorksheet(sheetName);
        return this;
    }

    public IStyleableStage InitializeByCollection<T>(
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

    public IStyleableStage BoldHeaders(string? headerInitializationCell = ExcelBuilderConstants.DefaultHeadersInitializationCell)
    {

        return this;
    }
    
    public IBuildStage AsBuildable() => this;

    public byte[] Build()
    {
        ThrowIfWorkbookIsDisposed();

        using var memoryStream = new MemoryStream();

        memoryStream.Seek(0, SeekOrigin.Begin);
        
        this._workbook.SaveAs(memoryStream);
        this._workbook.Dispose();

        return memoryStream.ToArray();
    }

    public void BuildAndSave(string path)
    {
        ThrowIfWorkbookIsDisposed();

        this._workbook.SaveAs(path);
        
        this._workbook.Dispose();
    }

    public void Dispose()
    {
        this._workbook.Dispose();
    }

    private void ThrowIfWorkbookIsDisposed()
    {
        if (_workbook == null)
            throw new InvalidOperationException("Workbook instance has been disposed already");
    }

}