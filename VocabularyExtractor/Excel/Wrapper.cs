using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text;

namespace VocabularyExtractor.Excel;

public static class Wrapper
{
    public static List<Vocabulary> ExtractVocabulary(
        Stream document,
        string firstCellValueStarter)
    {
        using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(document, false);

        if (spreadsheetDocument.WorkbookPart is null || spreadsheetDocument.WorkbookPart.Workbook.Sheets is null)
            throw new Exception("Workbook Part is null or there are no sheets in the document.");

        List<Vocabulary> result = new();
        foreach (Sheet sheet in spreadsheetDocument.WorkbookPart.Workbook.Sheets.Elements<Sheet>())
        {
            if (sheet.Id is null || sheet.Id.Value is null) continue;
            WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id.Value);

            Row[] rows = worksheetPart.Worksheet.Elements<SheetData>().First().Elements<Row>().ToArray();
            if (rows.Length < 1) continue;

            bool startExtraction = false;
            foreach (Row row in rows)
            {
                Cell[] cells = row.Elements<Cell>().ToArray();
                if (cells.Length < 0) continue;

                if (!startExtraction)
                {
                    string? firstCellValue = GetCellValue(spreadsheetDocument.WorkbookPart.SharedStringTablePart, cells[0]);
                    if (firstCellValue == firstCellValueStarter)
                        startExtraction = true;

                    continue;
                }

                int word, memoryAid, translation;
                switch (cells.Length)
                {
                    case 3:
                        (word, memoryAid, translation) = (0, 1, 2);
                        break;
                    case 5:
                        (word, memoryAid, translation) = (0, 2, 3);
                        break;
                    default:
                        continue;
                }

                result.Add(new(
                    GetCellValue(spreadsheetDocument.WorkbookPart.SharedStringTablePart, cells[word]),
                    GetCellValue(spreadsheetDocument.WorkbookPart.SharedStringTablePart, cells[memoryAid]),
                    GetCellValue(spreadsheetDocument.WorkbookPart.SharedStringTablePart, cells[translation])));
            }
        }

        return result;
    }

    static string? GetCellValue(
        SharedStringTablePart? stringTablePart,
        Cell cell) =>
        stringTablePart is not null && cell.DataType is not null && cell.DataType.Value == CellValues.SharedString && cell.CellValue is not null ?
        stringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cell.CellValue.InnerText)).InnerText :
        null;


    public static string ConvertVocabularyToLernbox(
        List<Vocabulary> list,
        string subject)
    {
        StringBuilder builder = new StringBuilder().AppendLine($"{subject}:");

        foreach (Vocabulary v in list)
            builder.AppendLine($"/ {v.Translation.Replace("/", " ∕ ")} / {v.Word.Replace("/", " ∕ ")} ; {v.MemoryAid.Replace("/", " ∕ ")}");

        return builder.ToString();
    }


    public static bool ValidateConfig(
        Config config)
    {
        if (string.IsNullOrEmpty(config.Subject) || string.IsNullOrEmpty(config.FirstCellValueStarter))
        {
            ConsoleHelpers.Write("Subject or First Cell Value Starter is empty. Please first update the config.");
            return false;
        }

        return true;
    }
}