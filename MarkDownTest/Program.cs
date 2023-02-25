using Markdig.OpenXml;

namespace MarkDownTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string md1 =
                "Dies ist ein test\r\n" +
                "ohne Markdown\r\n" +
                "\r\n" +
                "end";

            string md2 =
                "Dies ist ein test\r\n" +
                "# Head1\r\n" +
                "mit Markdown\r\n" +
                "\r\n" +
                "end";

            string md3 =
                "Dummy\r\n" +
                "# Head 1\r\n" +
                "Dummy\r\n" +
                "## Head 2\r\n" +
                "Dummy\r\n" +
                "### Head 3\r\n" +
                "Dummy\r\n" +
                "#### Head 4\r\n" +
                "Dummy\r\n" +
                "##### Head 5\r\n" +
                "Dummy\r\n" +
                "###### Head 6\r\n" +
                "Dummy\r\n";

            string md4 =
                "Dummy\r\n" +
                "* Item A1\r\n" +
                "* Item A2\r\n" +
                "\t* Item B1\r\n" +
                "\t* Item B2\r\n" +
                "\t\t* Item C1\r\n" +
                "\t\t* Item C2\r\n" +
                "\t\t\t* Item D1\r\n" +
                "\t\t\t* Item D2\r\n" +
                "\t\t\t\t* Item E1\r\n" +
                "\t\t\t\t* Item E2\r\n" +
                "\t\t\t\t\t* Item F1\r\n" +
                "\t\t\t\t\t* Item F2\r\n" +
                "Dummy\r\n";

            string md5 =
                "Dummy\r\n" +
                "1. Num A1\r\n" +
                "1. Num A2\r\n" +
                "\t1. Num B1\r\n" +
                "\t1. Num B2\r\n" +
                "\t\t1. Num C1\r\n" +
                "\t\t1. Num C2\r\n" +
                "\t\t\t1. Num D1\r\n" +
                "\t\t\t1. Num D2\r\n" +
                "\t\t\t\t1. Num E1\r\n" +
                "\t\t\t\t1. Num E2\r\n" +
                "\t\t\t\t\t1. Num F1\r\n" +
                "\t\t\t\t\t1. Num F2\r\n" +
                "Dummy\r\n";




            MarkdownOpenXmlConverter.CreateXmlDocument("C:\\Temp\\Test.docx", md3 + md4 + md5);
        }
    }
}