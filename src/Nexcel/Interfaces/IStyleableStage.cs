using Nexcel.Misc;

namespace Nexcel.Interfaces;

public interface IStyleableStage
{
    IStyleableStage BoldHeaders(string? headerInitializationCell = ExcelBuilderConstants.DefaultHeadersInitializationCell);
    IBuildStage AsBuildable();
}
