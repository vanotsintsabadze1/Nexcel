namespace Nexcel.Interfaces;

public interface IBuildStage
{
    byte[] Build();
    void BuildAndSave(string path);
}
