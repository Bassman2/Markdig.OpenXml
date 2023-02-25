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

            if (block.Parent is MarkdownDocument)
            {
            renderer.AddParagraph();
            }
            else if (block.Parent is ListItemBlock listItemBlock)
            {
                ListBlock listBlock = listItemBlock.Parent as ListBlock;

                int index = 1;
                Block b = listItemBlock;
                while (b.Parent is ListBlock)
                {
                    b = b.Parent.Parent;
                    index++;
                }
                
                renderer.AddListItem(index, listItemBlock.Order);
            }
            else throw new Exception();
                    
            renderer.WriteChildren(block.Inline);
            Debug.Unindent();
            Debug.WriteLine("--ParagraphBlock");
        }
    }
}
