using Markdig.Syntax;
using System.Diagnostics;

namespace Markdig.OpenXml.Renderer
{
    public class HeadingRenderer : OpenXmlObjectRenderer<HeadingBlock>
    {
        protected override void Write(OpenXmlRenderer renderer, HeadingBlock block)
        {
            Debug.WriteLine($"++HeadingBlock {block.Level}");
            Debug.Indent();

            renderer.AddHeading(block.Level);
            renderer.WriteChildren(block.Inline);
            Debug.Unindent();
            Debug.WriteLine("--HeadingBlock");
        }
    }
}
