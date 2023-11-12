namespace VocabularyExtractor.OnlineShare;

public static class Wrapper
{
    static readonly HttpClient httpClient = new();
    static readonly Random random = new();

    public static async Task<int> UploadAsync(
        string text,
        string apiKey)
    {
        int key = random.Next(10000, 99999);

        httpClient.DefaultRequestHeaders.Add("cl1papitoken", apiKey);
        HttpResponseMessage response = await httpClient.PostAsync($"https://api.cl1p.net/{key}", new StringContent(text));
        
        response.EnsureSuccessStatusCode();
        return key;
    }


    public static bool ValidateConfig(
        Config config)
    {
        if (string.IsNullOrEmpty(config.OnlineShareApiKey))
        {
            ConsoleHelpers.Write("Online Share API Key is empty. Please first update the config.");
            return false;
        }

        return true;
    }
}