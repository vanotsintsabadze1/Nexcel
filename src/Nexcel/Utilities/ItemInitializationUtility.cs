using ClosedXML.Excel;
using Nexcel.Misc;

namespace Nexcel.Utilities;

internal static class ItemInitializationUtility
{
    public static void InitializeWorksheetHeaders(
        IXLWorksheet worksheet, 
        IEnumerable<string> headers, 
        string? initializationStartCell = ExcelBuilderConstants.DefaultHeadersInitializationCell,
        bool? boldHeaders = true)
    {
        worksheet.Cell(initializationStartCell!).InsertData(new[] { headers });

        if (boldHeaders.HasValue && boldHeaders.Value)
        {
            var startCellCharacters = GetCellAddressLetter(initializationStartCell!);
            var startCellNumber = GetCellAddressNumber(initializationStartCell!);
            var nthLetterFromStartCell = NthColumnFrom(startCellCharacters, headers.Count());

            var range = worksheet.Range($"{initializationStartCell}:{nthLetterFromStartCell}{startCellNumber}");

            range.Style.Font.Bold = true;
        }
    }

    public static void InitializeWorksheetItems<T>(
        IXLWorksheet worksheet, 
        IEnumerable<T> items, 
        string? initializationStartCell = ExcelBuilderConstants.DefaultItemInitializationCell)
    {
        worksheet.Cell(initializationStartCell!).InsertData(items);
    }

    internal static string GetCellAddressLetter(string address)
    {
        char? firstNumericalCharacter = address.FirstOrDefault(char.IsDigit);
        int? indexOfFirstNumericalCharacter = firstNumericalCharacter != null ? address.IndexOf(firstNumericalCharacter.Value) : null;

        if (!indexOfFirstNumericalCharacter.HasValue) return address;
        return address.Remove(indexOfFirstNumericalCharacter.Value);
    }

    private static int GetCellAddressNumber(string address)
    {
        char? firstNumericalCharacter = address.FirstOrDefault(char.IsDigit);
        int? indexOfFirstNumericalCharacter = firstNumericalCharacter != null ? address.IndexOf(firstNumericalCharacter.Value) : null;

        if (!indexOfFirstNumericalCharacter.HasValue)
            throw new ArgumentException($"Given address {address} does not have any numerical representation");

        if (address.Length == 1 && indexOfFirstNumericalCharacter.HasValue)
            throw new ArgumentException($"Given address {address} does not have any letter representation");

        var addressTrimmedOfLetters = address.Remove(0, indexOfFirstNumericalCharacter.Value);

        var conversionWasValid = int.TryParse(addressTrimmedOfLetters, out var convertedAddress);

        if (!conversionWasValid) throw new ArgumentException($"Something went wrong while performing: {nameof(GetCellAddressNumber)} method");

        return convertedAddress;
    }

    public static int ColumnToIndex(string column)
    {
        int index = 0;
        for (int i = 0; i < column.Length; i++)
        {
            index *= 26;
            index += (column[i] - 'A' + 1);
        }
        return index - 1; // make zero-based
    }

    public static string IndexToColumn(int index)
    {
        string column = string.Empty;
        while (index >= 0)
        {
            column = (char)('A' + (index % 26)) + column;
            index = (index / 26) - 1;
        }
        return column;
    }
    public static string NthColumnFrom(string startColumn, int n)
    {
        int startIndex = ColumnToIndex(startColumn);
        int targetIndex = startIndex + n;
        if (targetIndex < 0)
            throw new ArgumentOutOfRangeException(nameof(n), "Resulting column is before A");
        return IndexToColumn(targetIndex);
    }
}
