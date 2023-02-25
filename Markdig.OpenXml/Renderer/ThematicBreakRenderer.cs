using Markdig.Syntax;

namespace Markdig.OpenXml.Renderer
{
    public class ThematicBreakRenderer : OpenXmlObjectRenderer<ThematicBreakBlock>
    {
        protected override void Write(OpenXmlRenderer renderer, ThematicBreakBlock block)
        {
            renderer.WriteChildren(block.Inline);
        }
    }
}
