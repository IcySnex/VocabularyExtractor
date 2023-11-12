using LovePdf.Core;
using LovePdf.Model.Task;

namespace VocabularyExtractor.LovePdf;

public static class Wrapper
{
    public static Task<byte[]> ConvertPdfToXlsxAsync(
        string publicKey,
        string secretKey,
        string filePath)
    {
        LovePdfApi api = new(publicKey, secretKey);

        PdfToOfficeTask task = api.CreateTask<PdfToOfficeTask>();
        task.AddFile(filePath);
        task.Process(new PdfToOfficeParams("xlsx"));

        return task.DownloadFileAsByteArrayAsync();
    }

    public static Task<byte[]> ConvertPdfToJpgAsync(
        string publicKey,
        string secretKey,
        string filePath)
    {
        LovePdfApi api = new(publicKey, secretKey);

        PdfToJpgTask task = api.CreateTask<PdfToJpgTask>();
        task.AddFile(filePath);
        task.Process();

        return task.DownloadFileAsByteArrayAsync();
    }


    public static bool ValidateConfig(
        Config config)
    {
        if (string.IsNullOrEmpty(config.ILovePDFPublicKey) || string.IsNullOrEmpty(config.ILovePDFSecretKey))
        {
            ConsoleHelpers.Write("ILovePDF Public Key or Secret Key is empty. Please first update the config.");
            return false;
        }

        return true;
    }
}