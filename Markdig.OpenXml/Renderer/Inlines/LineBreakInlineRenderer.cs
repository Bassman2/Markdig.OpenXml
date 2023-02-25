using Markdig.Syntax.Inlines;
using System.Diagnostics;

namespace Markdig.OpenXml.Renderer.Inlines
{
    public class LineBreakInlineRenderer : OpenXmlObjectRenderer<LineBreakInline>
    {
        protected override void Write(OpenXmlRenderer renderer, LineBreakInline obj)
        {
            Debug.WriteLine("**LineBreak");
            renderer.AddLineBreak();
        }
    }
}
