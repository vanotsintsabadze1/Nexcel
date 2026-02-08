namespace Nexcel.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class ExcelComplexMappingAttribute : Attribute
{
    internal string Cell { get; set; }
    internal int? HorizontallyMerges { get; set; }
    internal int? VerticallyMerges { get; set; }
    internal bool? Bold { get; set; }
    internal string? LocalizedValue { get; set; }

    public ExcelComplexMappingAttribute(
        string cell, 
        int? horizontallyMerges, 
        int? verticallyMerges,
        string? localizedValue,
        bool? bold = true)
    {
        Cell = cell;
        HorizontallyMerges = horizontallyMerges;
        VerticallyMerges = verticallyMerges;
        Bold = bold;
        LocalizedValue = localizedValue;
    }
}