using NUnit.Framework;
using OfficeOpenXml;
using System.Xml;

namespace EPPlusTest
{
    [TestFixture]
    public class ExcelStyleTests
    {
        [Test]
        public void QuotePrefixStyle()
        {
            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("QuotePrefixTest");
                var cell = ws.Cells["B2"];
                cell.Style.QuotePrefix = true;
                Assert.That(cell.Style.QuotePrefix);

                p.Workbook.Styles.UpdateXml();                
                var nodes = p.Workbook.StylesXml.SelectNodes("//d:cellXfs/d:xf", p.Workbook.NameSpaceManager);
                // Since the quotePrefix attribute is not part of the default style,
                // a new one should be created and referenced.
                Assert.That(0, Is.Not.EqualTo(cell.StyleID));
                Assert.That(nodes[0].Attributes["quotePrefix"], Is.Null);
                Assert.That("1", Is.EqualTo(nodes[cell.StyleID].Attributes["quotePrefix"].Value));
            }
        }
    }
}
