using Nexcel.Misc;

namespace Nexcel.Interfaces;

public interface IInitializedWorksheetStage : IDisposable
{
    IInitializedWorksheetStage InitializeByCollection<T>(
    IEnumerable<T> items,
    string? headerInitializationStartCell = ExcelBuilderConstants.DefaultHeadersInitializationCell,
    string? initializationStartCell = ExcelBuilderConstants.DefaultItemInitializationCell);

    IInitializedWorkbookStage Done();
}
