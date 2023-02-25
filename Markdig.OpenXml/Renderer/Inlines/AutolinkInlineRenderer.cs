using Markdig.Syntax.Inlines;
using System.Diagnostics;

namespace Markdig.OpenXml.Renderer.Inlines
{
    public class AutolinkInlineRenderer : OpenXmlObjectRenderer<AutolinkInline>
    {
        protected override void Write(OpenXmlRenderer renderer, AutolinkInline obj)
        {
            Debug.WriteLine($"++AutolinkInline {obj.Url}");
            Debug.Indent();
            //renderer.WriteChildren(block.Inline);
            Debug.Unindent();
            Debug.WriteLine("--AutolinkInline");
        }
    }
}
