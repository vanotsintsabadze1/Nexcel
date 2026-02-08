using ClosedXML;
using Nexcel.Utilities;

namespace Nexcel.Attributes;

[AttributeUsage(AttributeTargets.Property | AttributeTargets.Field)]
public class OrientedExcelComplexMappingAttribute : Attribute
{
    internal string Cell { get; }

    public OrientedExcelComplexMappingAttribute(Type typeToRefer, string nameOfProperty)
    {
        var property = typeToRefer.GetProperty(nameOfProperty);
        var positionOfReferredProperty = property.GetAttributes<ExcelComplexMappingAttribute>().FirstOrDefault();

        if (positionOfReferredProperty == null) throw new ArgumentException($"Property {nameOfProperty} does not have attribute {nameof(ExcelComplexMappingAttribute)}");

        Cell = ItemInitializationUtility.GetCellAddressLetter(positionOfReferredProperty.Cell);
    }
}