using Markdig.Syntax;
using System.Diagnostics;

namespace Markdig.OpenXml.Renderer
{
    public class ListRenderer : OpenXmlObjectRenderer<ListBlock>
    {
        protected override void Write(OpenXmlRenderer renderer, ListBlock block)
        {
            Debug.WriteLine($"++ListBlock {block.BulletType} {block.IsOrdered}");
            Debug.Indent();

            renderer.AddList(block.BulletType, block.IsOrdered, 0);

            foreach (var item in block) 
            {
                ListItemBlock listItem = item as ListItemBlock;
                Debug.WriteLine("++ListBlock.Item");
                Debug.Indent();

                
                renderer.Write(listItem);

                Debug.Unindent();
                Debug.WriteLine("--ListBlock.Item");
            }

            Debug.Unindent();
            Debug.WriteLine("--ListBlock");
        }
    }
}
