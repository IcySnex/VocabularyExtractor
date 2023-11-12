using Newtonsoft.Json;

namespace VocabularyExtractor;

public class Config
{
    public static Config Load(string path)
    {
        if (File.Exists(path))
        {
            string json = File.ReadAllText(path);
            return JsonConvert.DeserializeObject<Config>(json) ?? new();
        }

        return new();
    }

    public void Save(string path)
    {
        string json = JsonConvert.SerializeObject(this, Formatting.Indented);
        File.WriteAllText(path, json);
        Console.WriteLine("Configuration saved.");
    }

    public string ILovePDFPublicKey { get; set; } = string.Empty;

    public string ILovePDFSecretKey { get; set; } = string.Empty;

    public string OnlineShareApiKey { get; set; } = string.Empty;

    public string FirstCellValueStarter { get; set; } = string.Empty;

    public string Subject { get; set; } = string.Empty;
}