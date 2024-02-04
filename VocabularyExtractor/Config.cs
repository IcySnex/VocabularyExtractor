using Newtonsoft.Json;

namespace VocabularyExtractor;

public class Config
{
    public static Config Load(
        string path)
    {
        Thread.Sleep(1000);

        if (File.Exists(Path.Combine(AppContext.BaseDirectory, path)))
        {
            string json = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, path));
            return JsonConvert.DeserializeObject<Config>(json) ?? new();
        }

        return new();
    }

    public void Save(
        string path)
    {
        string json = JsonConvert.SerializeObject(this, Formatting.Indented);
        File.WriteAllText(Path.Combine(AppContext.BaseDirectory, path), json);
        Console.WriteLine("Configuration saved.");
    }

    public string ILovePDFPublicKey { get; set; } = string.Empty;

    public string ILovePDFSecretKey { get; set; } = string.Empty;

    public string OnlineShareApiKey { get; set; } = string.Empty;

    public string FirstCellValueStarter { get; set; } = string.Empty;

    public string Subject { get; set; } = string.Empty;

    public int? WordColumn { get; set; } = 0;
    
    public int? MemoryAidColumn { get; set; } = 1;

    public int? TranslationColumn { get; set; } = 2;
}