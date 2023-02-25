using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Markdig.OpenXml
{
    internal static class OpenXmlExtentions
    {
        #region Styles

        public static void AddParagraphStyle(this Styles styles, string styleId, bool isDefault, bool isCustomStyle, string name, bool isPrimaryStyle)
        {
            Style style = new Style() { Type = StyleValues.Paragraph,  StyleId = styleId, Default = isDefault, CustomStyle = isCustomStyle };
            style.Append(new StyleName() { Val = name });
            if (isPrimaryStyle)
            {
                style.Append(new PrimaryStyle());
            }
            styles.Append(style);
        }

        public static void AddParagraphStyle(this Styles styles, string styleId, bool isDefault, bool isCustomStyle, string name, string basedOn, string nextParagraphStyle, string linkedStyle, int uiPriority, bool isUnhideWhenUsed, bool isPrimaryStyle, string fontColor, int fontSize)
        {
            Style style = new Style() { Type = StyleValues.Paragraph, StyleId = styleId, Default = isDefault, CustomStyle = isCustomStyle };
            style.Append(new StyleName() { Val = name });

            style.Append(new BasedOn() { Val = basedOn });
            style.Append(new NextParagraphStyle() { Val = nextParagraphStyle });
            style.Append(new LinkedStyle() { Val = linkedStyle });
            style.Append(new UIPriority() { Val = uiPriority });

            if (isUnhideWhenUsed)
            {
                style.Append(new UnhideWhenUsed());
            }  

            if (isPrimaryStyle)
            {
                style.Append(new PrimaryStyle());
            }

            style.Append(new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new SpacingBetweenLines() { Before = "40", After = "0" },
                new OutlineLevel() { Val = 5 }
                ));

            style.Append(new StyleRunProperties(
                new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi },
                new Color () { Val = fontColor, ThemeShade = "BF", ThemeColor = ThemeColorValues.Accent1 },
                new FontSize () { Val = fontSize.ToString() },
                new FontSizeComplexScript() { Val = fontSize.ToString() }
                ));

            styles.Append(style);
        }

        public static void AddParagraphStyle(this Styles styles, OpenXmlStyle openXmlStyle)
            {
            Style style = new Style() { Type = StyleValues.Paragraph, StyleId = openXmlStyle.StyleId, CustomStyle = true };     // , Default = isDefault, CustomStyle = isCustomStyle
            style.Append(new StyleName() { Val = openXmlStyle.StyleId });

            style.Append(new BasedOn() { Val = "Normal" });
            style.Append(new NextParagraphStyle() { Val = "Normal" });
            style.Append(new LinkedStyle() { Val = openXmlStyle.StyleCharId });
            style.Append(new UIPriority() { Val = 9 });

            //if (isUnhideWhenUsed)
            //{
            //    style.Append(new UnhideWhenUsed());
            //}

            //if (isPrimaryStyle)
            //{
            //    style.Append(new PrimaryStyle());
            //}

            style.Append(new StyleParagraphProperties(
                new KeepNext(),
                new KeepLines(),
                new SpacingBetweenLines() { Before = "40", After = "0" },
                new OutlineLevel() { Val = 5 }
                ));

            style.Append(openXmlStyle.StyleRunProperties);

            styles.Append(style);

        }

        public static void AddParagraphStyle(this Styles styles, string styleId, bool isDefault, bool isCustomStyle, string name, string basedOn, int uiPriority, bool isPrimaryStyle)
        {
            Style style = new Style() { Type = StyleValues.Paragraph, StyleId = styleId, Default = isDefault, CustomStyle = isCustomStyle };
            style.Append(new StyleName() { Val = name });
            style.Append(new BasedOn() { Val = basedOn });
                style.Append(new UIPriority() { Val = uiPriority });
            if (isPrimaryStyle)
            {
                style.Append(new PrimaryStyle());
            }
            style.Append(new StyleParagraphProperties(
                new Indentation() { Left = "720" },
                new ContextualSpacing()
                )); 
            styles.Append(style);
            }

        public static void AddCharacterStyle(this Styles styles, string styleId, bool isDefault, bool isCustomStyle, string name, int uiPriority, bool isSemiHidden, bool isUnhideWhenUsed)
        {
            Style style = new Style() { Type = StyleValues.Character, StyleId = styleId, Default = isDefault, CustomStyle = isCustomStyle };
            style.Append(new StyleName() { Val = name });
            style.Append(new UIPriority() { Val = uiPriority });            
            if (isSemiHidden)
            {
                style.Append(new SemiHidden());
            }
            if (isUnhideWhenUsed)
            {
                style.Append(new UnhideWhenUsed());
            }
            styles.Append(style);
        }

        public static void AddCharacterStyle(this Styles styles, string styleId, bool isDefault, bool isCustomStyle, string name, string basedOn, string linkedStyle, int uiPriority, string fontColor, int fontSize)
            {
            Style style = new Style() { Type = StyleValues.Character, StyleId = styleId, Default = isDefault, CustomStyle = isCustomStyle };
            style.Append(new StyleName() { Val = name });
            style.Append(new BasedOn() { Val = basedOn });
            style.Append(new LinkedStyle() { Val = linkedStyle });
            style.Append(new UIPriority() { Val = uiPriority });
            style.Append(new StyleRunProperties(
                new RunFonts() { AsciiTheme = ThemeFontValues.MajorHighAnsi, HighAnsiTheme = ThemeFontValues.MajorHighAnsi, EastAsiaTheme = ThemeFontValues.MajorEastAsia, ComplexScriptTheme = ThemeFontValues.MajorBidi },
                new Color() { Val = fontColor, ThemeShade = "BF", ThemeColor = ThemeColorValues.Accent1 },
                new FontSize() { Val = fontSize.ToString() },
                new FontSizeComplexScript() { Val = fontSize.ToString() }
                ));
            styles.Append(style);
            }

        public static void AddCharacterStyle(this Styles styles, OpenXmlStyle openXmlStyle)
        {
            Style style = new Style() { Type = StyleValues.Character, StyleId = openXmlStyle.StyleCharId, CustomStyle = true };
            style.Append(new StyleName() { Val = openXmlStyle.StyleCharId });
            style.Append(new BasedOn() { Val = "DefaultParagraphFont" });
            style.Append(new LinkedStyle() { Val = openXmlStyle.StyleId });
            style.Append(new UIPriority() { Val = 9 });
            style.Append(openXmlStyle.StyleRunProperties);
            styles.Append(style);
        }

        public static void AddTableStyle(this Styles styles, string styleId, bool isDefault, bool isCustomStyle, string name, int uiPriority, bool isSemiHidden, bool isUnhideWhenUsed, int identation, int top, int left, int bottom, int right)
        {
            Style style = new Style() { Type = StyleValues.Table, StyleId = styleId, Default = isDefault, CustomStyle = isCustomStyle };
            style.Append(new StyleName() { Val = name });
            style.Append(new UIPriority() { Val = uiPriority });
            if (isSemiHidden)
            {
                style.Append(new SemiHidden());
            }
            if (isUnhideWhenUsed)
            {
                style.Append(new UnhideWhenUsed());
            }
            style.Append(new StyleTableProperties(
                new TableIndentation() { Width = identation, Type = TableWidthUnitValues.Dxa },
                new TableCellMarginDefault(
                    new TopMargin() { Width = top.ToString(), Type = TableWidthUnitValues.Dxa },
                    new TableCellLeftMargin() { Width = (short)left, Type = TableWidthValues.Dxa },
                    new BottomMargin() { Width = bottom.ToString(), Type = TableWidthUnitValues.Dxa },
                    new TableCellRightMargin() { Width = (short)right, Type = TableWidthValues.Dxa }
                    )
                ));
            styles.Append(style);
        }

        public static void AddNumberingStyle(this Styles styles, string styleId, bool isDefault, bool isCustomStyle, string name, int uiPriority, bool isSemiHidden, bool isUnhideWhenUsed)
        {
            Style style = new Style() { Type = StyleValues.Numbering, StyleId = styleId, Default = isDefault, CustomStyle = isCustomStyle };
            style.Append(new StyleName() { Val = name });
            style.Append(new UIPriority() { Val = uiPriority });
            if (isSemiHidden)
            {
                style.Append(new SemiHidden());
            }
            if (isUnhideWhenUsed)
            {
                style.Append(new UnhideWhenUsed());
            }
            styles.Append(style);
        }

        #endregion






    }
}
