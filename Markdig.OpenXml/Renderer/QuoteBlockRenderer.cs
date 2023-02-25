using Markdig.Syntax;

namespace Markdig.OpenXml.Renderer
{
    public class QuoteBlockRenderer : OpenXmlObjectRenderer<QuoteBlock>
    {
        protected override void Write(OpenXmlRenderer renderer, QuoteBlock block)
        {
            //renderer.WriteChildren(block.Inline);
        }
    }
}
