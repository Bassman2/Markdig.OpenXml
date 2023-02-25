using Markdig.Syntax;
using System.Diagnostics;

namespace Markdig.OpenXml.Renderer
{
    public class ParagraphRenderer : OpenXmlObjectRenderer<ParagraphBlock>
    {
        protected override void Write(OpenXmlRenderer renderer, ParagraphBlock block)
        {
            Debug.WriteLine("++ParagraphBlock");
            Debug.Indent();

            renderer.AddParagraph();
            renderer.WriteChildren(block.Inline);
            Debug.Unindent();
            Debug.WriteLine("--ParagraphBlock");
        }
    }
}
