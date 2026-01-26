using Nexcel.Builder;
using Nexcel.Misc;

namespace Nexcel.Interfaces;

public interface IInitializedWorkbookStage : IDisposable
{
    IInitializedWorksheetStage AddWorksheet(string? sheetName = "Sheet1");

    IInitializedWorkbookStage RemoveWorksheet(string? sheetName = null);

    IInitializedWorksheetStage SelectWorksheet(string sheetName);
    
    IInitializedWorksheetStage SelectWorksheet(int sheetIndex);

    byte[] Build();

    void BuildAndSave(string path);
}