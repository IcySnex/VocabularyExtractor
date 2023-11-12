using System.Diagnostics;

namespace VocabularyExtractor.Clipboard;

public class LinuxClipboard : IClipBoard
{
    public void SetText(
        string text)
    {
        ProcessStartInfo psi = new()
        {
            FileName = "xclip",
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