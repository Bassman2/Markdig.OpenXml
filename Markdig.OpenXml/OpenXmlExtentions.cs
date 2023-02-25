using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Markdig.OpenXml
{
    internal static class OpenXmlExtentions
    {
        public static void AddStyle(this Styles styles, string styleId, bool defaul, StyleValues type, string name, string basedOn = null, int uiPriority = 0, bool semiHidden = false, bool unhideWhenUsed = false, bool primaryStyle = false)
        {
            Style style = new Style() { StyleId = styleId, Default = defaul, Type = type };
            style.Append(new StyleName() { Val = name });

            if (basedOn != null)
            {
                style.Append(new BasedOn() { Val = "basedOn" });
            }  

            if (uiPriority > 0)
            {
                style.Append(new UIPriority() { Val = uiPriority });
            }

            if (semiHidden)
            {
                style.Append(new SemiHidden());
            }

            if (primaryStyle)
            {
                style.Append(new PrimaryStyle());
            }

            styles.Append(style);
        }

    }
}
