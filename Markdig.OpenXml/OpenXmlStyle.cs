using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Markdig.OpenXml
{
    public class OpenXmlStyle
    {
        public OpenXmlStyle(string styleId, string fontName, int fontSize, string fontColor)
        {
            this.FontName = fontName;
            this.FontSize = fontSize;
            this.FontColor = fontColor;
        }

        public string StyleId { get; set; }

        public string StyleCharId => StyleId + "Char";

        public string FontName { get; set; }

        public int FontSize { get; set; }

        public string FontColor { get; set; }

        internal RunProperties RunProperties =>
            new RunProperties(
                string.IsNullOrEmpty(this.FontName) 
                    ? new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi }  
                    : new RunFonts() { Ascii = FontName, HighAnsi = FontName, ComplexScript = FontName },
                new FontSize() { Val = FontSize.ToString() },
                new FontSizeComplexScript() { Val = FontSize.ToString() },
                new Color() { Val = FontColor }
                );

        internal StyleRunProperties StyleRunProperties =>
            new StyleRunProperties(
                string.IsNullOrEmpty(this.FontName)
                    ? new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi }
                    : new RunFonts() { Ascii = FontName, HighAnsi = FontName, ComplexScript = FontName },
                new FontSize() { Val = FontSize.ToString() },
                new FontSizeComplexScript() { Val = FontSize.ToString() },
                new Color() { Val = FontColor }
                );

        internal ParagraphProperties ParagraphProperties =>
            new ParagraphProperties(
                new ParagraphStyleId() { Val = StyleId }
                );

        //internal ParagraphProperties ListItemParagraphProperties =>
        //    new ParagraphProperties(
        //        new ParagraphStyleId() { Val = "ListParagraph" },
        //        new NumberingProperties(
        //            new NumberingLevelReference() { Val = 0 },
        //            new NumberingId() { Val = 1 }
        //            ));
    }
}
