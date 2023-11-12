using System.Runtime.InteropServices;
using System.Text;

namespace VocabularyExtractor.Clipboard;

public class WindowsClipboard : IClipBoard
{
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool OpenClipboard(IntPtr hWndNewOwner);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EmptyClipboard();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern IntPtr SetClipboardData(uint uFormat, IntPtr hMem);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool CloseClipboard();

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern IntPtr GlobalLock(IntPtr hMem);

    [DllImport("kernel32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GlobalUnlock(IntPtr hMem);


    private const uint CF_UNICODETEXT = 13;

    private const uint GMEM_MOVEABLE = 0x0002;


    public void SetText(string text)
    {
        if (!OpenClipboard(IntPtr.Zero))
            return;

        try
        {
            EmptyClipboard();
            byte[] data = Encoding.Unicode.GetBytes(text);
            IntPtr ptr = GlobalAlloc(GMEM_MOVEABLE, (UIntPtr)data.Length);
            if (ptr != IntPtr.Zero)
            {
                IntPtr lockedMem = GlobalLock(ptr);
                if (lockedMem == IntPtr.Zero)
                    return;

                Marshal.Copy(data, 0, lockedMem, data.Length);
                GlobalUnlock(lockedMem);
                SetClipboardData(CF_UNICODETEXT, ptr);
            }
        }
        finally
        {
            CloseClipboard();
        }
    }

}