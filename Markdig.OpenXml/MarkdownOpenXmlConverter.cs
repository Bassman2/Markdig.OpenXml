using DocumentFormat.OpenXml;
using Markdig.Syntax;
using System.Diagnostics;

namespace Markdig.OpenXml
{
    public static class MarkdownOpenXmlConverter
    {
        public static void AppendMarkdown(this OpenXmlCompositeElement parent, string markdown, OpenXmlFormat openXmlFormat = null)
        {
            Debug.WriteLine("#############################################################################");

            MarkdownDocument md = Markdig.Markdown.Parse(markdown);
            OpenXmlRenderer.Render(parent, md, openXmlFormat);
        }

        public static void CreateXmlDocument(string file, string markdown, OpenXmlFormat openXmlFormat = null)
        {
            Debug.WriteLine("#############################################################################");

            MarkdownDocument md = Markdig.Markdown.Parse(markdown);
            OpenXmlRenderer.Render(file, md, openXmlFormat);
        }
    }
    
   
}
