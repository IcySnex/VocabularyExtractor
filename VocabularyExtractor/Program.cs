namespace VocabularyExtractor;

public class Program
{
    static Config config = default!;
    static Clipboard.IClipBoard clipboard = Clipboard.IClipBoard.Default();

    public static async Task Main(string[] args)
    {
        config = Config.Load("config.json");

        if (args.Length > 0)
        {
            string option = args[0].ToLower();
            switch (option)
            {
                case "1":
                    if (!(args.Length > 2 && args[1].ToLower() == "-filepath"))
                    {
                        Console.WriteLine("Invalid parameters for extracting Vocabulary from PDF [1].\nUsage: VocabularyExtractor.exe 1 -filePath \"C:\\Some\\Example.pdf\"");
                        return;
                    }
                    await ExtractVocabularyAsync(args[2]);
                    break;
                case "2":
                    if (!(args.Length > 4 && args[1].ToLower() == "-filepath" && args[3].ToLower() == "-savetopath"))
                    {
                        Console.WriteLine("Invalid parameters for converting PDF to XLSX [2].\nUsage: VocabularyExtractor.exe 2 -filePath \"C:\\Some\\Example.pdf\" -saveToPath \"C:\\Save\\Example.xlsx\"");
                        return;
                    }
                    await ConvertToXlsxAsync(args[2], args[4]);
                    break;
                case "3":
                    if (!(args.Length > 4 && args[1].ToLower() == "-filepath" && args[3].ToLower() == "-savetopath"))
                    {
                        Console.WriteLine("Invalid parameters for converting PDF to JPG [3].\nUsage: VocabularyExtractor.exe 3 -filePath \"C:\\Some\\Example.pdf\" -saveToPath \"C:\\Save\\Example.jpg\"");
                        return;
                    }
                    await ConvertToJpgAsync(args[2], args[4]);
                    break;
                case "4":
                    if (!(args.Length > 2 && args[1].ToLower() == "-filepath"))
                    {
                        Console.WriteLine("Invalid parameters for saveing Vocabulary from spreadsheet [4].\nUsage: VocabularyExtractor.exe 4 -filePath \"C:\\Some\\Example.xlsx\" -copyToClipboard Y");
                        return;
                    }
                    bool copyToClipboard = args.Length > 4 && args[3].ToLower() == "-copytoclipboard" && args[4].ToLower() == "y";
                    SaveVocabulary(args[2], copyToClipboard);
                    break;
                case "5":
                    if (!(args.Length > 2 && args[1].ToLower() == "-text"))
                    {
                        Console.WriteLine("Invalid parameters for uploading Online Share Text [5].\nUsage: VocabularyExtractor.exe 5 -text \"Some random example\"");
                        return;
                    }
                    await UploadOnlineShare(args[2]);
                    break;
            }
            return;
        }

        while (true)
        {
            ConsoleHelpers.WriteClear("Vacabulary Sheet Extractor - Convert PDF tables to Excel tables and export vocabulary for Lernbox:\n");

            Console.WriteLine("[1]  Extract Vocabulary from PDF\n");
            Console.WriteLine("[2]  Convert PDF to XLSX");
            Console.WriteLine("[3]  Convert PDF to JPG");
            Console.WriteLine("[4]  Save Vocabulary from spreadsheet");
            Console.WriteLine("[5]  Upload Online Share Text\n");
            Console.WriteLine("[C]  Update Configuration");
            Console.WriteLine("[X]  Exit\n\n");

            string? choice = ConsoleHelpers.GetResponse("Press [1] - [6] to continue", "Invalid choice. Please enter a number between [1] and [5].");
            if (choice is null) continue;

            switch (choice.ToLower())
            {
                case "1":
                    await ExtractVocabularyAsync();
                    break;
                case "2":
                    await ConvertToXlsxAsync();
                    break;
                case "3":
                    await ConvertToJpgAsync();
                    break;
                case "4":
                    SaveVocabulary();
                    break;
                case "5":
                    await UploadOnlineShare();
                    break;
                case "c":
                    UpdateConfiguration(config);
                    break;
                case "x":
                    Exit();
                    return;
            }
        }
    }



    static void UpdateConfiguration(
        Config config)
    {
        while (true)
        {
            ConsoleHelpers.WriteClear("[C]  Update Configuration:\n");

            Console.WriteLine("[1] Change ILovePDF Public Key");
            Console.WriteLine("[2] Change ILovePDF Secret Key");
            Console.WriteLine("[3] Change Online Share API Key");
            Console.WriteLine("[4] Change First Cell Value Starter");
            Console.WriteLine("[5] Change Subject\n");
            Console.WriteLine("[X] Exit Configuration Update\n\n");

            string? choice = ConsoleHelpers.GetResponse("Press [1] - [6] to continue", "Invalid choice. Please enter a number between [1] and [5].");
            if (choice is null) continue;

            Console.Clear();
            switch (choice.ToLower())
            {
                case "1":
                    config.ILovePDFPublicKey = ConsoleHelpers.GetResponse("Enter new ILovePDF Public Key", null) ?? string.Empty;
                    Console.WriteLine("Public Key updated.");
                    break;
                case "2":
                    config.ILovePDFSecretKey = ConsoleHelpers.GetResponse("Enter new ILovePDF Secret Key", null) ?? string.Empty;
                    Console.WriteLine("Secret Key updated.");
                    break;
                case "3":
                    config.OnlineShareApiKey = ConsoleHelpers.GetResponse("Enter new Online Share API Key", null) ?? string.Empty;
                    Console.WriteLine("Online Share API Key updated.");
                    break;
                case "4":
                    config.FirstCellValueStarter = ConsoleHelpers.GetResponse("Enter new First Cell Value Starter", null) ?? string.Empty;
                    Console.WriteLine("First Cell Value Starter updated.");
                    break;
                case "5":
                    config.Subject = ConsoleHelpers.GetResponse("Enter new Subject", null) ?? string.Empty;
                    Console.WriteLine("Subject updated.");
                    break;
                case "x":
                    config.Save("config.json");
                    ConsoleHelpers.Write("\nConfiguration updated.");
                    return;
            }
        }
    }

    static void Exit()
    {
        ConsoleHelpers.WriteClear("[X]  Exit:");
        Environment.Exit(0);
    }


    static async Task ExtractVocabularyAsync(
        string? filePath = null)
    {
        ConsoleHelpers.WriteClear("[1]  Extract Vocabulary from PDF:\n");

        if (!LovePdf.Wrapper.ValidateConfig(config)) return;
        if (!Excel.Wrapper.ValidateConfig(config)) return;
        if (!OnlineShare.Wrapper.ValidateConfig(config)) return;

        if (filePath is null)
            filePath = ConsoleHelpers.GetResponse("Enter the file path to your PDF", "File path can not be empty.");
        if (filePath is null) return;

        try
        {
            Console.Clear();
            Console.WriteLine("Converting PDF to XLSX...");
            byte[] file = await LovePdf.Wrapper.ConvertPdfToXlsxAsync(config.ILovePDFPublicKey, config.ILovePDFSecretKey, filePath);

            Console.WriteLine("Saving Vocabulary from spreadsheet...");
            using Stream document = new MemoryStream(file);
            List<Excel.Vocabulary> result = Excel.Wrapper.ExtractVocabulary(document, config.FirstCellValueStarter);
            if (result is null || result.Count < 1)
            {
                ConsoleHelpers.Write($"Saving Vocabulary failed. For some reason there are no vocabulary in the list?");
                return;
            }
            string lernboxResult = Excel.Wrapper.ConvertVocabularyToLernbox(result, config.Subject);

            Console.WriteLine("Uploading to Online Share...");
            int id = await OnlineShare.Wrapper.UploadAsync(lernboxResult, config.OnlineShareApiKey);

            ConsoleHelpers.Write($"\nExtracted Vocabulary from PDF successfuly.\nYou can copy your Lernbox vocabulary on: https://cl1p.net/{id}.");
        }
        catch (Exception ex)
        {
            ConsoleHelpers.Write($"Extracting Vocabulary from PDF (Exception: {ex.Message}).");
        }

    }

    static async Task ConvertToXlsxAsync(
        string? filePath = null,
        string? saveToPath = null)
    {
        ConsoleHelpers.WriteClear("[2]  Convert PDF to XLSX:\n");

        if (!LovePdf.Wrapper.ValidateConfig(config)) return;

        if (filePath is null)
            filePath = ConsoleHelpers.GetResponse("Enter the file path to your PDF", "File path can not be empty.");
        if (filePath is null) return;

        if (saveToPath is null)
            saveToPath = ConsoleHelpers.GetResponse("Enter the file path to which you want to save the XLSX", "File path can not be empty.");
        if (saveToPath is null) return;

        try
        {
            ConsoleHelpers.WriteClear("Converting PDF to XLSX...");
            byte[] file = await LovePdf.Wrapper.ConvertPdfToXlsxAsync(config.ILovePDFPublicKey, config.ILovePDFSecretKey, filePath);

            ConsoleHelpers.WriteClear("PDF finished converting to XLSX.\nSaving to XLSX now...");
            File.WriteAllBytes(saveToPath, file);

            ConsoleHelpers.Write("\nXLSX saved successfuly.");
        }
        catch (Exception ex)
        {
            ConsoleHelpers.Write($"Converting PDF failed (Exception: {ex.Message}).");
        }
    }

    static async Task ConvertToJpgAsync(
        string? filePath = null,
        string? saveToPath = null)
    {
        ConsoleHelpers.WriteClear("[3]  Convert PDF to JPG:\n");

        if (!LovePdf.Wrapper.ValidateConfig(config)) return;

        if (filePath is null)
            filePath = ConsoleHelpers.GetResponse("Enter the file path to your PDF", "File path can not be empty.");
        if (filePath is null) return;

        if (saveToPath is null)
            saveToPath = ConsoleHelpers.GetResponse("Enter the file path to which you want to save the jpg (if more than on page it will be saved as a zip)", "File path can not be empty.");
        if (saveToPath is null) return;

        try
        {
            ConsoleHelpers.WriteClear("Converting PDF to JPG...");
            byte[] file = await LovePdf.Wrapper.ConvertPdfToJpgAsync(config.ILovePDFPublicKey, config.ILovePDFSecretKey, filePath);

            ConsoleHelpers.WriteClear("PDF finished converting to JPG.\nSaving to JPG/Archive now...");
            File.WriteAllBytes(saveToPath, file);

            ConsoleHelpers.Write("\nJPG/Archive saved successfuly.");
        }
        catch (Exception ex)
        {
            ConsoleHelpers.Write($"Converting PDF failed (Exception: {ex.Message}).");
        }
    }

    static void SaveVocabulary(
        string? filePath = null,
        bool? copyToClipboard = null)
    {
        ConsoleHelpers.WriteClear("[4]  Save Vocabulary from spreadsheet:\n");

        if (!Excel.Wrapper.ValidateConfig(config)) return;

        if (filePath is null)
            filePath = ConsoleHelpers.GetResponse("Enter the file path to your spreadsheet", "File path can not be empty.");
        if (filePath is null) return;

        try
        {
            ConsoleHelpers.WriteClear("Saving Vocabulary from spreadsheet...");
            using Stream document = File.OpenRead(filePath);
            List<Excel.Vocabulary> result = Excel.Wrapper.ExtractVocabulary(document, config.FirstCellValueStarter);
            if (result is null || result.Count < 1)
            {
                ConsoleHelpers.Write($"Saving Vocabulary failed. For some reason there are no vocabulary in the list?");
                return;
            }
            string lernboxResult = Excel.Wrapper.ConvertVocabularyToLernbox(result, config.Subject);

            if (copyToClipboard is null)
                copyToClipboard = ConsoleHelpers.AskResponse("Vocabulary finished savinf.\nPress [Y] to copy the result to your clipboard or press [N] to write it to the console.");
            if (copyToClipboard.Value)
            {
                Console.WriteLine("Copying content to clipboard...");
                clipboard.SetText(lernboxResult);
            }
            else
            {

                Console.WriteLine("Writing result to config...\n");
                Console.WriteLine(lernboxResult);
            }

            ConsoleHelpers.Write($"\nExtracted Vocabulary successfuly.\nFirst word is: \"{result[0].Word}\", Last word is: \"{result[result.Count - 1].Word}\".");
        }
        catch (Exception ex)
        {
            ConsoleHelpers.Write($"Extracting Vocabulary failed (Exception: {ex.Message}).");
        }
    }

    static async Task UploadOnlineShare(
        string? text = null)
    {
        ConsoleHelpers.WriteClear("[5]  Upload Online Share Text:\n");

        if (!OnlineShare.Wrapper.ValidateConfig(config)) return;

        if (text is null)
            text = ConsoleHelpers.GetResponse("Enter the text you want to upload to Online Share", "Text can not be empty.");
        if (text is null) return;

        try
        {
            ConsoleHelpers.WriteClear("Uploading to Online Share...");
            int id = await OnlineShare.Wrapper.UploadAsync(text, config.OnlineShareApiKey);

            ConsoleHelpers.Write($"\nUploaded to Online Share.\nYou can visit your text by: https://cl1p.net/{id}.");
        }
        catch (Exception ex)
        {
            ConsoleHelpers.Write($"Uploading to Online Share failed (Exception: {ex.Message}).");
        }
    }
}