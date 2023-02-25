using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Markdig.OpenXml
{
    public class OpenXmlStyle
    {
        public OpenXmlStyle(string fontName, int fontSize, string fontColor)
        {
            this.FontName = fontName;
            this.FontSize = fontSize;
            this.FontColor = fontColor;
        }

        public string FontName { get; set; }

        public int FontSize { get; set; }

        public string FontColor { get; set; }

        internal RunProperties RunProperties =>
            new RunProperties(
                new RunFonts() { Ascii = FontName, HighAnsi = FontName, ComplexScript = FontName },
                new FontSize() { Val = FontSize.ToString() },
                new FontSizeComplexScript() { Val = FontSize.ToString() },
                new Color() { Val = FontColor }
                );
    }
}
