using Nexcel.Misc;

namespace Nexcel.Interfaces;

public interface IItemInitializationStage
{
    IStyleableStage InitializeByCollection<T>(
        IEnumerable<T> items,
        string? headerInitializationStartCell = ExcelBuilderConstants.DefaultHeadersInitializationCell, 
        string? initializationStartCell = ExcelBuilderConstants.DefaultItemInitializationCell);

    IBuildStage AsBuildable();
}