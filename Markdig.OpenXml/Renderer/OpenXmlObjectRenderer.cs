using Markdig.Renderers;
using Markdig.Syntax;

namespace Markdig.OpenXml.Renderer
{
    public abstract class OpenXmlObjectRenderer<TObject> : MarkdownObjectRenderer<OpenXmlRenderer, TObject> where TObject : MarkdownObject
    {
    }
}
