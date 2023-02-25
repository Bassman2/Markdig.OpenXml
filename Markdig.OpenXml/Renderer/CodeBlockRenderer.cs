using Markdig.Syntax;
using System;

namespace Markdig.OpenXml.Renderer
{
    public class CodeBlockRenderer : OpenXmlObjectRenderer<CodeBlock>
    {
        protected override void Write(OpenXmlRenderer renderer, CodeBlock block)
        {
            renderer.WriteChildren(block.Inline);
        }
    }
}
