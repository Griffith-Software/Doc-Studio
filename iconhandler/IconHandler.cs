// Doc Studio icon handler: a small File Explorer add-on that gives files opened with Doc Studio an icon for their
// type (resources\filetypes\<ext>.ico: the Doc Studio page with a coloured band naming the type).
//
// Windows 11 records "open with Doc Studio" as the whole app, which has one icon. Explorer asks this handler
// instead (Applications\Doc Studio.exe\DefaultIcon = "%1" with shellex\IconHandler = this class), and it answers
// with the icon file for the file's extension. It does nothing else: no file is opened or read, only its name is
// looked at. Built with the C# compiler that comes with Windows (.NET Framework 4); see tools/make-icon-handler.js.
// Doc Studio registers it for the current Windows user (no admin) and removes it when uninstalled.

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

[assembly: System.Reflection.AssemblyTitle("Doc Studio icon handler")]
[assembly: System.Reflection.AssemblyCompany("Griffith Software")]
[assembly: System.Reflection.AssemblyProduct("Doc Studio")]
[assembly: System.Reflection.AssemblyCopyright("Copyright (c) 2026 Richard Griffith (d/b/a Griffith Software)")]
[assembly: System.Reflection.AssemblyVersion("1.0.0.0")]

namespace DocStudio
{
    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("000214FA-0000-0000-C000-000000000046")]
    public interface IExtractIconW
    {
        [PreserveSig]
        int GetIconLocation(uint uFlags, [MarshalAs(UnmanagedType.LPWStr, SizeParamIndex = 2)] StringBuilder szIconFile, int cchMax, out int piIndex, out uint pwFlags);
        [PreserveSig]
        int Extract([MarshalAs(UnmanagedType.LPWStr)] string pszFile, uint nIconIndex, IntPtr phiconLarge, IntPtr phiconSmall, uint nIconSize);
    }

    [ComVisible(true), Guid("6C1E2B7A-4F3D-4D8E-9B21-5A7C0D3E9F41"), ClassInterface(ClassInterfaceType.None)]
    public class IconHandler : IPersistFile, IExtractIconW
    {
        const int S_OK = 0, S_FALSE = 1, E_FAIL = unchecked((int)0x80004005);
        string ext = "";

        // IPersist / IPersistFile: Explorer gives the file's name; only its extension is kept.
        public void GetClassID(out Guid pClassID) { pClassID = typeof(IconHandler).GUID; }
        public int IsDirty() { return S_FALSE; }
        public void Load(string pszFileName, int dwMode)
        {
            try { ext = (Path.GetExtension(pszFileName ?? "") ?? "").TrimStart('.').ToLowerInvariant(); } catch { ext = ""; }
        }
        public void Save(string pszFileName, bool fRemember) { }
        public void SaveCompleted(string pszFileName) { }
        public void GetCurFile(out string ppszFileName) { ppszFileName = null; }

        // IExtractIconW: the icon file for this extension; Explorer extracts it itself.
        public int GetIconLocation(uint uFlags, StringBuilder szIconFile, int cchMax, out int piIndex, out uint pwFlags)
        {
            piIndex = 0;
            pwFlags = 0;
            try
            {
                string ico = IconFor(ext);
                if (ico == null || szIconFile == null || ico.Length >= cchMax) return S_FALSE;
                szIconFile.Clear();
                szIconFile.Append(ico);
                return S_OK;
            }
            catch { return S_FALSE; }
        }
        public int Extract(string pszFile, uint nIconIndex, IntPtr phiconLarge, IntPtr phiconSmall, uint nIconSize) { return S_FALSE; }

        // <Doc Studio>\resources\filetypes\<ext>.ico. The folder is written next to this file by Doc Studio
        // (filetypes.txt), so the handler follows Doc Studio wherever it is installed.
        static string folder;
        static string IconFor(string e)
        {
            if (string.IsNullOrEmpty(e) || e.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) return null;
            if (folder == null)
            {
                string here = Path.GetDirectoryName(typeof(IconHandler).Assembly.Location);
                string pointer = Path.Combine(here, "filetypes.txt");
                folder = File.Exists(pointer) ? File.ReadAllText(pointer).Trim() : "";
            }
            if (folder.Length == 0) return null;
            string ico = Path.Combine(folder, e + ".ico");
            return File.Exists(ico) ? ico : null;
        }
    }
}
