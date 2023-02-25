using Markdig.Syntax.Inlines;
using System.Diagnostics;

namespace Markdig.OpenXml.Renderer.Inlines
{
    public class EmphasisInlineRenderer : OpenXmlObjectRenderer<EmphasisInline>
    {
        protected override void Write(OpenXmlRenderer renderer, EmphasisInline obj)
        {
            Debug.WriteLine("++EmphasisInline");
            Debug.Indent();
            foreach (Inline item in obj)
            {
                Debug.WriteLine($"++EmphasisInline.Item {obj.DelimiterChar} {obj.DelimiterCount}");
                Debug.Indent();

                if (item is LiteralInline literal)
                { 
                    Debug.WriteLine(literal.Content.ToString());
                }
                else if (item is EmphasisInline emphasis)
                {
                    renderer.WriteChildren(emphasis);
                }
                else
                {
                    throw new Exception();
                }

                Debug.Unindent();
                Debug.WriteLine("--EmphasisInline.Item");
            }

            Debug.Unindent();
            Debug.WriteLine("--EmphasisInline");
        }
    }
}
