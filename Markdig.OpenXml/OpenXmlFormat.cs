namespace Markdig.OpenXml
{
    public class OpenXmlFormat
    {
        private OpenXmlFormat()
        {
            
        }

        public static OpenXmlFormat Document
        {
            get
            {
                OpenXmlFormat format = new();
                format.Language = "en-US";
                format.TextStyle = new OpenXmlStyle("Normal", "Arial", 16, "000000");
                format.Header1Style = new OpenXmlStyle("Heading1", "Arial", 24, "000000");
                format.Header2Style = new OpenXmlStyle("Heading2", "Arial", 22, "000000");
                format.Header3Style = new OpenXmlStyle("Heading3", "Arial", 20, "000000");
                format.Header4Style = new OpenXmlStyle("Heading4", "Arial", 19, "000000");
                format.Header5Style = new OpenXmlStyle("Heading5", "Arial", 18, "000000");
                format.Header6Style = new OpenXmlStyle("Heading6", "Arial", 17, "000000");
                format.ListStyle = new OpenXmlStyle("ListParagraph", "Arial", 18, "000000");
                

                //format.LinkStyle = new OpenXmlStyle("Arial", 16, "0563C1");
                return format;
            }
        }


        public static OpenXmlFormat Inline
        {
            get
            {
                OpenXmlFormat format = new();
                format.Language = "en-US";
                format.TextStyle = new OpenXmlStyle("MDNormal", "Arial", 16, "000000");
                format.Header1Style = new OpenXmlStyle("MDHeading1", "Arial", 24, "000000");
                format.Header2Style = new OpenXmlStyle("MDHeading2", "Arial", 22, "000000");
                format.Header3Style = new OpenXmlStyle("MDHeading3", "Arial", 20, "000000");
                format.Header4Style = new OpenXmlStyle("MDHeading4", "Arial", 19, "000000");
                format.Header5Style = new OpenXmlStyle("MDHeading5", "Arial", 18, "000000");
                format.Header6Style = new OpenXmlStyle("MDHeading6", "Arial", 17, "000000");
                format.ListStyle = new OpenXmlStyle("ListParagraph", "Arial", 18, "000000");

                //format.LinkStyle = new OpenXmlStyle("Arial", 16, "0563C1");
                return format;
            }
        }

        public string Language { get; set; }

        public OpenXmlStyle TextStyle { get; set; }

        public OpenXmlStyle Header1Style { get; set; }
        public OpenXmlStyle Header2Style { get; set; }
        public OpenXmlStyle Header3Style { get; set; }
        public OpenXmlStyle Header4Style { get; set; }
        public OpenXmlStyle Header5Style { get; set; }
        public OpenXmlStyle Header6Style { get; set; }
        public OpenXmlStyle ListStyle { get; set; }
        public OpenXmlStyle LinkStyle { get; set; }


        internal OpenXmlStyle[] HeaderStyles => new OpenXmlStyle[] { null, Header1Style, Header2Style, Header3Style, Header4Style, Header5Style, Header6Style };
        


        //public string ParagraphStyleName { get; set; }

        //public string HeadingStyleName { get; set; }

        //public string RunStyleText { get; set; }

        //public string RunStyleLink { get; set; }

        //public string FontName { get; set; }

        //public int FontSize { get; set; }

        //public string TextColor { get; set; }
        //public string LinkColor { get; set; }
    }
}
