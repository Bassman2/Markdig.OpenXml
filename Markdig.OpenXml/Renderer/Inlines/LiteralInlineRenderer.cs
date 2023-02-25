using Markdig.Syntax.Inlines;
using System.Diagnostics;

namespace Markdig.OpenXml.Renderer.Inlines
{
    public class LiteralInlineRenderer : OpenXmlObjectRenderer<LiteralInline>
    {
        protected override void Write(OpenXmlRenderer renderer, LiteralInline obj)
        {
            Debug.WriteLine("++LiteralInline");
            Debug.Indent();

            string s = obj.Content.ToString();
            Debug.WriteLine(s);
            renderer.AddText(s);

            Debug.Unindent();
            Debug.WriteLine("--LiteralInline");
        }
    }
}
