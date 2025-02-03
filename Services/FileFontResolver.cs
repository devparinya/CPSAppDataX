using PdfSharp.Fonts;

namespace QueueAppManager.Service
{
    public class FileFontResolver : IFontResolver // FontResolverBase
    {
        public string DefaultFontName => throw new NotImplementedException();

        public byte[] GetFont(string faceName)
        {
            using (var ms = new MemoryStream())
            {
                using (var fs = File.Open(faceName, FileMode.Open))
                {
                    fs.CopyTo(ms);
                    ms.Position = 0;
                    return ms.ToArray();
                }
            }
        }

        public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            if (familyName.Equals("thsarabun", StringComparison.CurrentCultureIgnoreCase))
            {
                if (isBold && isItalic)
                {
                    return new FontResolverInfo("fonts/thsarabun-bold-italic.ttf");
                }
                else if (isBold)
                {
                    return new FontResolverInfo("fonts/thsarabun-bold.ttf");
                }
                else if (isItalic)
                {
                    return new FontResolverInfo("fonts/thsarabun-italic.ttf");
                }
                else
                {
                    return new FontResolverInfo("fonts/thsarabun.ttf");
                }
            }
            return null;
        }
    }
}
