using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using Markdig.OpenXml.Renderer;
using Markdig.OpenXml.Renderer.Inlines;
using Markdig.Renderers;
using Markdig.Syntax;
using System.Xml.Linq;


namespace Markdig.OpenXml
{
    public class OpenXmlRenderer : RendererBase
    {
        public static void Render(OpenXmlCompositeElement parent, MarkdownDocument md, OpenXmlFormat format = null)
        {
            new OpenXmlRenderer().Render(parent, md, format);
        }

        public static void Render(string file, MarkdownDocument md, OpenXmlFormat format = null)
        {
            new OpenXmlRenderer().Render(file, md, format);
        }

        private OpenXmlFormat format;
        private OpenXmlCompositeElement parent;
        private Paragraph paragraph;
        //private OpenXmlStyle style;
        //private Run run;

        public OpenXmlRenderer()
        {
            ObjectRenderers.Add(new CodeBlockRenderer());
            ObjectRenderers.Add(new ListRenderer());
            ObjectRenderers.Add(new HeadingRenderer());
            ObjectRenderers.Add(new HtmlBlockRenderer());
            ObjectRenderers.Add(new ParagraphRenderer());
            ObjectRenderers.Add(new QuoteBlockRenderer());
            ObjectRenderers.Add(new ThematicBreakRenderer());

            //// Default inline renderers
            ObjectRenderers.Add(new AutolinkInlineRenderer());
            ObjectRenderers.Add(new CodeInlineRenderer());
            ObjectRenderers.Add(new DelimiterInlineRenderer());
            ObjectRenderers.Add(new EmphasisInlineRenderer());
            ObjectRenderers.Add(new LineBreakInlineRenderer());
            ObjectRenderers.Add(new HtmlInlineRenderer());
            ObjectRenderers.Add(new HtmlEntityInlineRenderer());
            ObjectRenderers.Add(new LinkInlineRenderer());
            ObjectRenderers.Add(new LiteralInlineRenderer());
        }

        public void Render(OpenXmlCompositeElement parent, MarkdownObject markdownObject, OpenXmlFormat format = null)
        {
            this.format = format ?? OpenXmlFormat.Inline;
            this.parent = parent;
            Render(markdownObject);
        }

        public void Render(string file, MarkdownObject markdownObject, OpenXmlFormat format = null)
        {
            this.format = format ?? OpenXmlFormat.Document;
            using (WordprocessingDocument wordDocument = WordprocessingDocument.Create(file, WordprocessingDocumentType.Document))
            {
                // Add a main document part. 
                MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                SetStyleDefinitions(mainPart);
                SetNumberingDefinitions(mainPart);

                // Create the document structure and add some text.
                mainPart.Document = new();
                Body body = mainPart.Document.AppendChild(new Body());


                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text("Create text in body - CreateWordprocessingDocument"));

                parent = body;
                Render(markdownObject);
            }
        }

        private void SetStyleDefinitions(MainDocumentPart mainPart)
        {
            StyleDefinitionsPart styleDefinitionsPart = mainPart.StyleDefinitionsPart ?? mainPart.AddNewPart<StyleDefinitionsPart>();
            Styles styles = styleDefinitionsPart.Styles ??= new();

            styles.Append(new DocDefaults(
                new RunPropertiesDefault(
                    new RunPropertiesBaseStyle(
                        new RunFonts() { AsciiTheme = ThemeFontValues.MinorHighAnsi, HighAnsiTheme = ThemeFontValues.MinorHighAnsi, EastAsiaTheme = ThemeFontValues.MinorHighAnsi, ComplexScriptTheme = ThemeFontValues.MinorBidi },
                        new FontSize() { Val = "22" },
                        new FontSizeComplexScript() { Val = "22" },
                        new Languages() { Val = this.format.Language, EastAsia = "en-US", Bidi = "ar-SA" }
                    )
                ),
                new ParagraphPropertiesDefault
                (
                    new ParagraphPropertiesBaseStyle(
                        new SpacingBetweenLines() { After = "160", Line = "259", LineRule = LineSpacingRuleValues.Auto }
                    )
                )
            ));

            styles.AddParagraphStyle("Normal",               true, false, "Normal", true);
            styles.AddCharacterStyle("DefaultParagraphFont", true, false, "Default Paragraph Font", 1, true, true);
            styles.AddTableStyle    ("TableNormal",          true, false, "Normal Table", 99, true, true, 0, 0, 180, 0, 108);
            styles.AddNumberingStyle("NoList",               true, false, "Normal", 99, true, true);

            //styles.AddParagraphStyle("Heading1", false, false, "heading 1", "Normal", "Normal", "Heading1Char", 9, false, false, "2F5496", 32);
            //styles.AddParagraphStyle("Heading2", false, false, "heading 2", "Normal", "Normal", "Heading2Char", 9, true, true, "2F5496", 26);
            //styles.AddParagraphStyle("Heading3", false, false, "heading 3", "Normal", "Normal", "Heading3Char", 9, true, true, "2F5496", 24);
            //styles.AddParagraphStyle("Heading4", false, false, "heading 4", "Normal", "Normal", "Heading4Char", 9, true, true, "2F5496", 22);
            //styles.AddParagraphStyle("Heading5", false, false, "heading 5", "Normal", "Normal", "Heading5Char", 9, true, true, "2F5496", 22);
            //styles.AddParagraphStyle("Heading6", false, false, "heading 6", "Normal", "Normal", "Heading6Char", 9, true, true, "2F5496", 22);

            //styles.AddCharacterStyle("Heading1Char", false, true, "Heading 1 Char", "DefaultParagraphFont", "Heading1", 9, "2F5496", 32);
            //styles.AddCharacterStyle("Heading2Char", false, true, "Heading 2 Char", "DefaultParagraphFont", "Heading2", 9, "2F5496", 26);
            //styles.AddCharacterStyle("Heading3Char", false, true, "Heading 3 Char", "DefaultParagraphFont", "Heading3", 9, "2F5496", 24);
            //styles.AddCharacterStyle("Heading4Char", false, true, "Heading 4 Char", "DefaultParagraphFont", "Heading4", 9, "2F5496", 22);
            //styles.AddCharacterStyle("Heading5Char", false, true, "Heading 5 Char", "DefaultParagraphFont", "Heading5", 9, "2F5496", 22);
            //styles.AddCharacterStyle("Heading6Char", false, true, "Heading 6 Char", "DefaultParagraphFont", "Heading6", 9, "2F5496", 22);

            styles.AddParagraphStyle(this.format.Header1Style);
            styles.AddParagraphStyle(this.format.Header2Style);
            styles.AddParagraphStyle(this.format.Header3Style);
            styles.AddParagraphStyle(this.format.Header4Style);
            styles.AddParagraphStyle(this.format.Header5Style);
            styles.AddParagraphStyle(this.format.Header6Style);

            styles.AddCharacterStyle(this.format.Header1Style);
            styles.AddCharacterStyle(this.format.Header2Style);
            styles.AddCharacterStyle(this.format.Header3Style);
            styles.AddCharacterStyle(this.format.Header4Style);
            styles.AddCharacterStyle(this.format.Header5Style);
            styles.AddCharacterStyle(this.format.Header6Style);

            //
            // ListParagraph
            //
            styles.Append(new Style(
                new StyleName() { Val = "List Paragraph" },
                new BasedOn() { Val = "Normal" },
                new UIPriority() { Val = 34 },
                new PrimaryStyle(),
                new Rsid() { Val = "009951DC" },      // ?????????????????
                new StyleParagraphProperties(
                    new Indentation() { Left = "720" },
                    new ContextualSpacing()
                ))
                { Type = StyleValues.Paragraph, StyleId = "ListParagraph" }
            );
        }

        private void SetNumberingDefinitions(MainDocumentPart mainPart)
        {
            NumberingDefinitionsPart numberingDefinitionsPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
            Numbering numbering = numberingDefinitionsPart.Numbering ??= new();

            AbstractNum abstractNum = numbering.AppendChild(new AbstractNum());
            abstractNum.Append(new MultiLevelType() { Val = MultiLevelValues.HybridMultilevel });
            abstractNum.Append(new TemplateCode() { Val = "BA968A3E" } );

            for (int index = 0; index < 6; index++)
            {
                Level level = abstractNum.AppendChild(new Level() { LevelIndex = index, TemplateCode = "04070003", Tentative = true });
                level.Append(new StartNumberingValue() { Val = 1 });
                level.Append(new NumberingFormat() { Val = NumberFormatValues.Bullet });
                level.Append(new LevelText() { Val = "" });
                level.Append(new LevelJustification() { Val = LevelJustificationValues.Left });
                level.Append(new PreviousParagraphProperties(new Indentation() { Left = (720 * (index + 1)).ToString() , Hanging = "360" }));
                level.Append(new NumberingSymbolRunProperties(new RunFonts() { Ascii = "Symbol", HighAnsi = "Symbol", Hint = FontTypeHintValues.Default }));
            }


            NumberingInstance numberingInstance = numbering.AppendChild(new NumberingInstance() { NumberID = 1 });
            numberingInstance.Append(new AbstractNumId() { Val = 0 });

        }

        public override object Render(MarkdownObject markdownObject)
        {
            
            Write(markdownObject);
            return parent;
        }

        public void AddParagraph()
        {
            //this.style = this.format.TextStyle;
            this.paragraph = this.parent.AppendChild(new Paragraph(this.format.TextStyle.ParagraphProperties));
        }

        public void AddListItem(int level, int id)
        {
            this.paragraph = this.parent.AppendChild(
                new Paragraph(
                    new ParagraphProperties(
                        new ParagraphStyleId() { Val = "ListParagraph" },
                        new NumberingProperties(
                            new NumberingLevelReference() { Val = level },
                            new NumberingId() { Val = id }
                        ))));

        }

        public void AddHeading(int level)
        {
            //this.style = this.format.HeaderStyles[level];
            this.paragraph = this.parent.AppendChild(new Paragraph(this.format.HeaderStyles[level].ParagraphProperties));
        }

        public void AddList(char bulletType, bool isOrdered, int indent)
        {
            //AddParagraph(format.ParagraphStyleName, format.RunStyleText, format.FontName, format.FontSize, format.TextColor);
            //ParagraphProperties paragraphProperties = this.paragraph.GetFirstChild<ParagraphProperties>();

            //if (isOrdered)
            //{
            //    NumberingProperties numberingProperties = paragraphProperties.AppendChild(new NumberingProperties());
            //    numberingProperties.AppendChild(new NumberingLevelReference() { Val = indent });
            //    numberingProperties.AppendChild(new NumberingId() { Val = 38 });
            //}
            //else
            //{
            //    NumberingProperties numberingProperties = paragraphProperties.AppendChild(new NumberingProperties());
            //    numberingProperties.AppendChild(new NumberingLevelReference() { Val = indent });
            //    numberingProperties.AppendChild(new NumberingId() { Val = 37 });
            //}

        }

        //public void AddParagraph()
        //{
        //    this.style = this.format.TextStyle;

        //    this.paragraph = this.parent.AppendChild(new Paragraph());
        //}

        //public Run AddRun(OpenXmlStyle style)
        //{
        //    Run run = this.paragraph.AppendChild(new Run());
        //    //RunProperties runProperties = run.AppendChild(style.RunProperties);
        //    return run;
        //}

        public void AddText(string text)
        {
            Run run = this.paragraph.AppendChild(new Run());
            run.AppendChild(new Text() { Text = text });
        }

        public void AddLineBreak()
        {
            this.paragraph.AppendChild(new Run(new Break()));
        }
    }
}


/*
   public class MarkdownConverter
   {
       private readonly string colorText = "000000";  // black
       private readonly string colorLink = "0563C1";  
       private readonly OpenXmlCompositeElement parent;
       private readonly string font = "Arial";
       private readonly int fontSize = 16;
       private readonly MainDocumentPart mainDocumentPart;


       //private readonly string paragraphStyle = "HTMLVorformatiert";
       private readonly string runStyleText = "Nachrichtenkopfbeschriftun";
       private readonly string runStyleLink = "Hyperlink";
       //private readonly string runStyleHead1 = "berschrift1";
       //private readonly string runStyleHead2 = "berschrift2";
       //private readonly string runStyleHead3 = "berschrift3";
       //private readonly string runStyleHead4 = "berschrift4";
       //private readonly string runStyleHead5 = "berschrift5";
       //private readonly string runStyleHead6 = "berschrift6";

       private Paragraph currentParagraph;

       public MarkdownConverter(OpenXmlCompositeElement parent, string font = "Arial", int fontSize = 16)
       { 
           this.parent = parent;
           this.font = font;
           this.fontSize = fontSize;
           this.mainDocumentPart = parent.Ancestors<Document>().Single().MainDocumentPart;
       }

       public void ParseMarkdown(string txt)
       {
           const char NewLine = '\n';
           // convert line endings to UNIX LF
           txt = txt.Replace("\r\n", "\n").Replace("\r", "\n");

           //ParserState parserState = ParserState.Text;

           StringBuilder text = new();

           LinkMode linkMode = LinkMode.Text;
           string linkTitle = String.Empty;
           string linkUri;

           bool backslash = false;
           int spaces = 0;
           int newLines = 1;   // start with one new line

           NewMode newMode = NewMode.None;

           // Paragraph
           ParagraphStyle currParagraphStyle = ParagraphStyle.Text;
           ParagraphStyle nextPrargraphStyle = ParagraphStyle.Text;

           // Run
           Emphasis emphasis = Emphasis.None;
           EmphasisMode emphasisMode = EmphasisMode.None;

           MDDebug($"<{txt}>");

           foreach (char c in txt)
           {
               MDDebug($"####### text='{text}' backslash={backslash} newLines={newLines} spaces={spaces} newMode={newMode} currParagraphStyle={currParagraphStyle} nextPrargraphStyle={nextPrargraphStyle} emphasisMode={emphasisMode} emphasis={emphasis} linkMode={linkMode}");

               MDDebug($"--<{c}>--");
               switch (c)
               {
               // Backslash
               case '\\' when !backslash:
                   backslash = true;
                   continue;

               case NewLine:
                   newLines++;
                   if (spaces >= 2)
                   {
                       newMode = NewMode.Line;
                   }
                   if (newLines > 1)
                   {
                       newMode = NewMode.Paragraph;
                   }
                   // End Heading
                   if (nextPrargraphStyle >= ParagraphStyle.Header1 && nextPrargraphStyle <= ParagraphStyle.Header6)
                   {
                       newMode = NewMode.Paragraph;
                       nextPrargraphStyle = ParagraphStyle.Text;
                   }
                   continue;

               // Emphasis
               case '*' when !backslash:
               case '_' when !backslash:
                   switch (emphasisMode)
                   {
                   case EmphasisMode.None:
                       emphasisMode = EmphasisMode.Begin;
                       emphasis = Emphasis.Italic;
                       continue;
                   case EmphasisMode.Begin:
                       emphasis = (Emphasis)Math.Min((int)emphasis + 1, (int)Emphasis.BoldAndItalic); 
                       newMode = NewMode.Run;
                       continue;
                   case EmphasisMode.End:
                       emphasis = Emphasis.None;
                       newMode = NewMode.Run;
                       continue;
                   }
                   continue;

               // Monospace
               case '\'' when !backslash:
                   newMode = NewMode.Paragraph;
                   nextPrargraphStyle = nextPrargraphStyle != ParagraphStyle.Monospace ? ParagraphStyle.Monospace : ParagraphStyle.Text;
                   continue;

               // Heading
               case '#' when newLines > 0 && !backslash:
                   nextPrargraphStyle = nextPrargraphStyle switch
                   {
                       ParagraphStyle.Header1 => ParagraphStyle.Header2,
                       ParagraphStyle.Header2 => ParagraphStyle.Header3,
                       ParagraphStyle.Header3 => ParagraphStyle.Header4,
                       ParagraphStyle.Header4 => ParagraphStyle.Header5,
                       ParagraphStyle.Header5 => ParagraphStyle.Header6,
                       ParagraphStyle.Header6 => ParagraphStyle.Header6,
                       _ => ParagraphStyle.Header1
                   };
                   newMode = NewMode.Paragraph;
                   continue;

               // Link    
               case '[' when linkMode != LinkMode.Text:
                   linkMode = LinkMode.Title;
                   newMode = NewMode.Run;
                   continue;
               case ']' when linkMode == LinkMode.Title:
                   linkMode = LinkMode.Between;
                   linkTitle = text.ToString();
                   text.Clear();
                   continue;
               case '(' when linkMode == LinkMode.Between:
                   linkMode = LinkMode.Link;
                   continue;
               case ')' when linkMode == LinkMode.Link:
                   linkMode = LinkMode.Text;
                   linkUri = text.ToString();
                   text.Clear();
                   if (linkUri.StartsWith("http"))
                   {
                       AddHyperlink(linkTitle, linkUri, emphasis);
                   }
                   else // not a link handle as text
                   {
                       text.Append($"[{linkTitle}]({linkUri}");
                       break;
                   }
                   continue;
               }

               // link break if ']' is not followed by '('
               if (linkMode == LinkMode.Between)
               {
                   linkMode = LinkMode.Text;
                   text.Append($"[{linkTitle}]");
               }

               // space counter
               spaces = c == ' ' ? spaces + 1 : 0;

               // start new Run
               if (emphasisMode == EmphasisMode.Begin)
               {
                   emphasisMode = EmphasisMode.Active;
               }
               if (emphasisMode == EmphasisMode.End)
               {
                   emphasisMode = EmphasisMode.None;
               }

               // create new Paragraph and Run instance
               switch (newMode)
               {
               case NewMode.Line:
                   AddLinebreak();
                   AddText(text, emphasis);
                   break;
               case NewMode.Run:
                   AddText(text, emphasis);
                   break;
               case NewMode.Paragraph:
                   if (!string.IsNullOrWhiteSpace(text.ToString()))
                   {
                       AddParagraph(currParagraphStyle);
                       AddText(text, emphasis);
                   }
                   currParagraphStyle = nextPrargraphStyle;
                   break;
               }
               newMode = NewMode.None;               

               text.Append(c);

               newLines = 0;
               backslash = false;

           }

           // append last text
           if (!string.IsNullOrEmpty(text.ToString()))
           {
               switch (newMode)
               {
               case NewMode.Line:
                   AddLinebreak();
                   AddText(text, emphasis);
                   break;
               case NewMode.Run:
                   AddText(text, emphasis);
                   break;
               case NewMode.Paragraph:
                   AddParagraph(currParagraphStyle);
                   AddText(text, emphasis);
                   break;
               }
           }
       }

       private enum LinkMode
       {
           Text,
           Title,
           Between,
           Link
       }

       private enum ParagraphStyle
       {
           Text,
           Header1,
           Header2,
           Header3,
           Header4,
           Header5,
           Header6,
           Enumeration1,
           Enumeration2,
           Enumeration3,
           Enumeration4,
           Enumeration5,
           Enumeration6,
           Numeration1,
           Numeration2,
           Numeration3,
           Numeration4,
           Numeration5,
           Numeration6,
           Monospace,
       }

       private enum NewMode
       {
           None,
           Line,
           Run, 
           Paragraph
       }

       [Flags]
       private enum Emphasis
       {
           None = 0,
           Italic = 1,
           Bold = 2,
           BoldAndItalic = 3
       }

       private enum EmphasisMode
       {
           None,
           Begin,
           Active,
           End
       }

       private void AddParagraph(ParagraphStyle paragraphStyle)
       {
           MDDebug($"Paragraph Style='{paragraphStyle}'");

           this.currentParagraph = this.parent.AppendChild(new Paragraph());

           ParagraphProperties paragraphProperties = this.currentParagraph.AppendChild(new ParagraphProperties());

           ParagraphStyleId paragraphStyleId = paragraphProperties.AppendChild(new ParagraphStyleId() 
           { 
               Val = paragraphStyle switch
               {
                   ParagraphStyle.Text => "HTMLVorformatiert",
                   ParagraphStyle.Header1 => "berschrift1",
                   ParagraphStyle.Header2 => "berschrift2",
                   ParagraphStyle.Header3 => "berschrift3",
                   ParagraphStyle.Header4 => "berschrift4",
                   ParagraphStyle.Header5 => "berschrift5",
                   ParagraphStyle.Header6 => "berschrift6",
                   _ => "HTMLVorformatiert"
               }

           });

           if (paragraphStyle >= ParagraphStyle.Enumeration1 && paragraphStyle <= ParagraphStyle.Enumeration6)
           {
               NumberingProperties numberingProperties = paragraphProperties.AppendChild(new NumberingProperties());
               numberingProperties.AppendChild(new NumberingLevelReference() { Val = paragraphStyle - ParagraphStyle.Enumeration1 });
               numberingProperties.AppendChild(new NumberingId() { Val = 37 });
           }
           if (paragraphStyle >= ParagraphStyle.Numeration1 && paragraphStyle <= ParagraphStyle.Numeration6)
           {
               NumberingProperties numberingProperties = paragraphProperties.AppendChild(new NumberingProperties());
               numberingProperties.AppendChild(new NumberingLevelReference() { Val = paragraphStyle - ParagraphStyle.Numeration1 });
               numberingProperties.AppendChild(new NumberingId() { Val = 38 });
           }

           Tabs tabs = paragraphProperties.AppendChild(new Tabs());
           tabs.AppendChild(new TabStop() { Position = 916, Val = TabStopValues.Clear });
           tabs.AppendChild(new TabStop() { Position = 1832, Val = TabStopValues.Clear });
           tabs.AppendChild(new TabStop() { Position = 2748, Val = TabStopValues.Clear });
           tabs.AppendChild(new TabStop() { Position = 3664, Val = TabStopValues.Clear });
           tabs.AppendChild(new TabStop() { Position = 4580, Val = TabStopValues.Clear });
           tabs.AppendChild(new TabStop() { Position = 5496, Val = TabStopValues.Clear });
           tabs.AppendChild(new TabStop() { Position = 6412, Val = TabStopValues.Clear });
           tabs.AppendChild(new TabStop() { Position = 7328, Val = TabStopValues.Clear });
           tabs.AppendChild(new TabStop() { Position = 8244, Val = TabStopValues.Clear });
           tabs.AppendChild(new TabStop() { Position = 9160, Val = TabStopValues.Clear });
           tabs.AppendChild(new TabStop() { Position = 10076, Val = TabStopValues.Clear });
           tabs.AppendChild(new TabStop() { Position = 100, Val = TabStopValues.Left });
           tabs.AppendChild(new TabStop() { Position = 2300, Val = TabStopValues.Left });

           paragraphProperties.AppendChild(new SpacingBetweenLines() { After = "160" });

           ParagraphMarkRunProperties paragraphMarkRunProperties = paragraphProperties.AppendChild(new ParagraphMarkRunProperties());

           paragraphMarkRunProperties.AppendChild(new RunStyle() { Val = this.runStyleText });
           paragraphMarkRunProperties.AppendChild(new RunFonts() { Ascii = font, HighAnsi = font, ComplexScript = font });
           paragraphMarkRunProperties.AppendChild(new Color() { Val = colorText });

           if (paragraphStyle >= ParagraphStyle.Header1 && paragraphStyle <= ParagraphStyle.Header6)
           {
               paragraphMarkRunProperties.AppendChild(new FontSize() { Val = (fontSize + (ParagraphStyle.Header6 - paragraphStyle + 1) * 4).ToString() });
           }
           else
           {
               paragraphMarkRunProperties.AppendChild(new FontSize() { Val = fontSize.ToString() });
           }
       }

       private void AddLinebreak()
       {
           MDDebug("  LineBreak");
           this.currentParagraph.Elements<Run>().LastOrDefault()?.AppendChild(new Break());
       }

       private void AddText(StringBuilder str, Emphasis emphasis)
       {
           string text = str.ToString();
           // clear for next "Run"
           str.Clear();

           MDDebug($"  Text='{text}'");

           if (string.IsNullOrEmpty(text))
           {
               return;
           }

           Run run = this.currentParagraph.AppendChild(new Run());

           RunProperties runProperties = run.AppendChild(new RunProperties());
           runProperties.RunStyle = new RunStyle() { Val = "Nachrichtenkopfbeschriftun" };
           runProperties.RunFonts = new RunFonts() { Ascii = this.font };
           runProperties.FontSize = new FontSize() { Val = this.fontSize.ToString() };
           //runProperties.FontSizeComplexScript = runProp.FontSizeComplexScript > 0 ? new FontSizeComplexScript() { Val = runProp.FontSizeComplexScript.ToString() } : null;
           runProperties.Color = new Color() { Val = new StringValue(this.colorText) };
           runProperties.Bold = new Bold() { Val = emphasis.HasFlag(Emphasis.Bold) };
           runProperties.Italic = new Italic() { Val = emphasis.HasFlag(Emphasis.Italic) };

           run.AppendChild(new Text() { Text = text });
       }

       private void AddText(OpenXmlCompositeElement element, string text, Emphasis emphasis)
       {

       }

       private void AddHyperlink(string title, string link, Emphasis emphasis)
       {
           MDDebug($"  Hyperlink Title='{title}' Link='{link}'");

           HyperlinkRelationship hyperlinkRelationship = this.mainDocumentPart.AddHyperlinkRelationship(new Uri(link), true);

           Hyperlink hyperlink = this.currentParagraph.AppendChild(new Hyperlink() { Id = hyperlinkRelationship.Id });

           Run run = hyperlink.AppendChild(new Run());

           RunProperties runProperties = run.AppendChild(new RunProperties());
           runProperties.RunStyle = new RunStyle() { Val = this.runStyleLink };
           runProperties.RunFonts = new RunFonts() { Ascii = this.font };
           runProperties.FontSize = new FontSize() { Val = this.fontSize.ToString() };
           //runProperties.FontSizeComplexScript = runProp.FontSizeComplexScript > 0 ? new FontSizeComplexScript() { Val = runProp.FontSizeComplexScript.ToString() } : null;
           runProperties.Color = new Color() { Val = this.colorLink };
           runProperties.Bold = new Bold() { Val = emphasis.HasFlag(Emphasis.Bold) }; ;
           runProperties.Italic = new Italic() { Val = emphasis.HasFlag(Emphasis.Italic) };

           run.AppendChild(new Text() { Text = title });
       }

       private static void MDDebug(string str)
       {
           Debug.WriteLine(str);
       }
   */

//this.paragraph = this.parent.AppendChild(new Paragraph());
//this.run = this.paragraph.AppendChild(new Run());

//RunProperties runProperties = run.AppendChild(new RunProperties());
//runProperties.RunStyle = new RunStyle() { Val = "Nachrichtenkopfbeschriftun" };
//runProperties.RunFonts = new RunFonts() { Ascii = this.font };
//runProperties.FontSize = new FontSize() { Val = this.fontSize.ToString() };
////runProperties.FontSizeComplexScript = runProp.FontSizeComplexScript > 0 ? new FontSizeComplexScript() { Val = runProp.FontSizeComplexScript.ToString() } : null;
//runProperties.Color = new Color() { Val = new StringValue(this.colorText) };
//runProperties.Bold = new Bold() { Val = emphasis.HasFlag(Emphasis.Bold) };
//runProperties.Italic = new Italic() { Val = emphasis.HasFlag(Emphasis.Italic) };

