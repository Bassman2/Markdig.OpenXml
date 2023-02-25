namespace Markdig.OpenXml
{
    public class OpenXmlFormat
    {
        public OpenXmlFormat()
        {
            this.Language = "en-US";
            this.TextStyle = new OpenXmlStyle("Arial", 16, "000000");
            this.Header1Style = new OpenXmlStyle("Arial", 24, "000000");
            this.Header2Style = new OpenXmlStyle("Arial", 22, "000000");
            this.Header3Style = new OpenXmlStyle("Arial", 20, "000000");
            this.Header4Style = new OpenXmlStyle("Arial", 19, "000000");
            this.Header5Style = new OpenXmlStyle("Arial", 18, "000000");
            this.Header6Style = new OpenXmlStyle("Arial", 17, "000000");
            this.LinkStyle = new OpenXmlStyle("Arial", 16, "0563C1");
        }

        public string Language { get; set; }

        public OpenXmlStyle TextStyle { get; set; }

        public OpenXmlStyle Header1Style { get; set; }
        public OpenXmlStyle Header2Style { get; set; }
        public OpenXmlStyle Header3Style { get; set; }
        public OpenXmlStyle Header4Style { get; set; }
        public OpenXmlStyle Header5Style { get; set; }
        public OpenXmlStyle Header6Style { get; set; }

        public OpenXmlStyle[] HeaderStyles => new OpenXmlStyle[] { null, Header1Style, Header2Style, Header3Style, Header4Style, Header5Style, Header6Style };

        public OpenXmlStyle LinkStyle { get; set; }


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
