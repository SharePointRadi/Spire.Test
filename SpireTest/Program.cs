using Spire.Xls;
using System.Xml;
using System;

namespace SpireTest
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello, Spire support team!");
            Console.WriteLine(DateTime.Now);
            string filePath = @"Excel Tutorial for Windows.xlsx";

            var path = Path.Combine(Environment.CurrentDirectory, filePath);
            try
            {
                var result = ConvertFileAsync(path).Result;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            //at System.Enum.TryParseByName(RuntimeType enumType, ReadOnlySpan`1 value, Boolean ignoreCase, Boolean throwOnFailure, UInt64 & result)
            //at System.Enum.TryParseInt32Enum(RuntimeType enumType, ReadOnlySpan`1 value, Int32 minInclusive, Int32 maxInclusive, Boolean ignoreCase, Boolean throwOnFailure, TypeCode type, Int32 & result)
            //at System.Enum.TryParse(Type enumType, ReadOnlySpan`1 value, Boolean ignoreCase, Boolean throwOnFailure, Object & result)
            //at System.Enum.Parse(Type enumType, String value, Boolean ignoreCase)
            //at spr겗.䛱(XmlTextReader A_0, Int32 & A_1)
            //at spr겗.䠖(XmlTextReader A_0, Int32 & A_1)
            //at spr겗.䠖(XmlTextReader A_0)
            //at spr겗.䤻(XmlTextReader A_0)
            //at spr겗.䩠(XmlTextReader A_0)
            //at spr겗.ᯙ(XmlTextReader A_0)
            //at spr뷂.䛱(spr䘂 A_0, Boolean A_1)
            //at spr뷂.䠖(XmlElement A_0, spr䘂 A_1, Boolean A_2)
            //at spr뷂.䛱(XmlElement A_0, spr䘂 A_1, Int32 A_2, Boolean A_3)
            //at spr뷂.䛱()
            //at spr퀒.䛱(spr䡌 A_0)
            //at spr퀒.䠖(spr䡌 A_0, String A_1, Boolean A_2)
            //at spr퀒.䩠(spr䡌 A_0, String A_1)
            //at spr퀒.䮅()
            //at spr퀒.曽()
            //at spr컭.䛱(spr᷺ A_0)
            //at spr᷺.䛱(Stream A_0, spr滐 A_1, Boolean A_2)
            //at spr᷺.䛱(String A_0, spr滐 A_1)
            //at spr᷺..ctor(String A_0, spr滐 A_1)
            //at Spire.Xls.Core.Spreadsheet.XlsWorkbook..ctor(Object A_0, String A_1, ExcelVersion A_2)
            //at Spire.Xls.Workbook.LoadFromFile(String fileName, ExcelVersion version)
            //at Spire.Xls.Workbook.LoadFromFile(String fileName)
            //at SpireTest.Program.ConvertFileAsync(String filePath) in C:\Projects\temp\SpireTest\SpireTest\Program.cs:line 19
            //at SpireTest.Program.Main(String[] args) in C:\Projects\temp\SpireTest\SpireTest\Program.cs:line 11
        }

        public static Task<Stream> ConvertFileAsync(string filePath)
        {
            var workbook = new Spire.Xls.Workbook();
            var outputStream = new MemoryStream();

            workbook.LoadFromFile(filePath);
            workbook.SaveToStream(outputStream, Spire.Xls.FileFormat.PDF);

            return Task.FromResult<Stream>(outputStream);
        }
    }
}