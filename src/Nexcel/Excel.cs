using ClosedXML.Excel;
using Nexcel.Builder;
using Nexcel.Interfaces;

namespace Nexcel;

public static class Excel
{
    public static IInitializedWorkbookStage CreateWorkbook()
    {
        var xlWorkbook = new XLWorkbook();
        return new ExcelBuilder(xlWorkbook);
    }
}