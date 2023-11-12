using System.Diagnostics;

namespace VocabularyExtractor.Clipboard;

public class MacOSClipboard : IClipBoard
{
    public void SetText(string text)
    {
        ProcessStartInfo psi = new()
        {
            FileName = "pbcopy",
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using Process process = new() { StartInfo = psi };
        process.Start();

        using StreamWriter sw = process.StandardInput;
        if (!sw.BaseStream.CanWrite)
            return;

        sw.Write(text);
    }
}