namespace Nexcel.Interfaces;

public interface IExcelComplexLayout
{
    IExcelComplexLayout MapHeaders<T>() where T : class;
}

