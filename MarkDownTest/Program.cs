using Markdig.OpenXml;

namespace MarkDownTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string md1 =
                "Dies ist ein test\r\n" +
                "ohne Markdown\r\r" +
                "\r\n" +
                "end";

            string md2 =
                "Dies ist ein test\r\n" +
                "# Head1\r\n" +
                "mit Markdown\r\r" +
                "\r\n" +
                "end";



            MarkdownOpenXmlConverter.CreateXmlDocument("C:\\Temp\\Test.docx", md2);
        }
    }
}