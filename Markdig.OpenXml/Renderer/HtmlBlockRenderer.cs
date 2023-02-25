using Markdig.Syntax;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Markdig.OpenXml.Renderer
{
    internal class HtmlBlockRenderer : OpenXmlObjectRenderer<HeadingBlock>
    {
        protected override void Write(OpenXmlRenderer renderer, HeadingBlock headingBlock)
        {
            renderer.WriteChildren(headingBlock.Inline);
        }
    }
}
