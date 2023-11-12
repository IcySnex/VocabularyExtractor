namespace VocabularyExtractor.Excel;

public class Vocabulary
{
    public Vocabulary(
        string? word,
        string? memoryAid,
        string? translation)
    {
        Word = word ?? string.Empty;
        MemoryAid = memoryAid ?? string.Empty;
        Translation = translation ?? string.Empty;
    }


    public string Word { get; set; }

    public string MemoryAid { get; set; }

    public string Translation { get; set; }
}