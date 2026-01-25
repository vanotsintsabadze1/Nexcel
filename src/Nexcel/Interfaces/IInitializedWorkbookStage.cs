using Nexcel.Builder;

namespace Nexcel.Interfaces;

public interface IInitializedWorkbookStage
{
    public IItemInitializationStage AddWorksheet(string? sheetName = "Sheet1");
}