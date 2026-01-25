using Nexcel;

var collection = new List<Data>() { new Data("hehe", "buehehe") };
Excel.CreateWorkbook().AddWorksheet().InitializeByCollection(collection).AsBuildable().BuildAndSave("C:\\Users\\tsint\\OneDrive\\Desktop\\test.xlsx");

public record Data(string Something, string Something2);