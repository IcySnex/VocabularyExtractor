using System.Runtime.InteropServices;

namespace VocabularyExtractor.Clipboard;

public interface IClipBoard
{
    void SetText(
        string text);


    public static IClipBoard Default()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            return new WindowsClipboard();
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
            return new LinuxClipboard();
        else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
            return new MacOSClipboard();
        else
            throw new NotSupportedException();
    }
}