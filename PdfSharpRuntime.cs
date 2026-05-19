using PdfSharp.Fonts;

namespace MinsPDFViewer
{
    internal static class PdfSharpRuntime
    {
        private static readonly object SyncRoot = new();
        private static bool _initialized;

        public static void EnsureInitialized()
        {
            if (_initialized)
                return;

            lock (SyncRoot)
            {
                if (_initialized)
                    return;

                if (GlobalFontSettings.FontResolver == null)
                    GlobalFontSettings.FontResolver = new WindowsFontResolver();

                _initialized = true;
            }
        }
    }
}
