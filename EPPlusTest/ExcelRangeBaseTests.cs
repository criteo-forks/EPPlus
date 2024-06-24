using NUnit.Framework;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestFixture]
    public class ExcelRangeBaseTests : TestBase
    {
        [Test]
        public void CopyCopiesCommentsFromSingleCellRanges()
        {
            InitBase();
            var pck = new ExcelPackage();
            var ws1 = pck.Workbook.Worksheets.Add("CommentCopying");
            var sourceExcelRange = ws1.Cells[3, 3];
            Assert.That(sourceExcelRange.Comment, Is.Null);
            sourceExcelRange.AddComment("Testing comment 1", "test1");
            Assert.That("test1", Is.EqualTo(sourceExcelRange.Comment.Author));
            Assert.That("Testing comment 1", Is.EqualTo(sourceExcelRange.Comment.Text));
            var destinationExcelRange = ws1.Cells[5, 5];
            Assert.That(destinationExcelRange.Comment, Is.Null);
            sourceExcelRange.Copy(destinationExcelRange);
            // Assert the original comment is intact.
            Assert.That("test1", Is.EqualTo(sourceExcelRange.Comment.Author));
            Assert.That("Testing comment 1", Is.EqualTo(sourceExcelRange.Comment.Text));
            // Assert the comment was copied.
            Assert.That("test1", Is.EqualTo(destinationExcelRange.Comment.Author));
            Assert.That("Testing comment 1", Is.EqualTo(destinationExcelRange.Comment.Text));
        }

        [Test]
        public void CopyCopiesCommentsFromMultiCellRanges()
        {
            InitBase();
            var pck = new ExcelPackage();
            var ws1 = pck.Workbook.Worksheets.Add("CommentCopying");
            var sourceExcelRangeC3 = ws1.Cells[3, 3];
            var sourceExcelRangeD3 = ws1.Cells[3, 4];
            var sourceExcelRangeE3 = ws1.Cells[3, 5];
            Assert.That(sourceExcelRangeC3.Comment, Is.Null);
            Assert.That(sourceExcelRangeD3.Comment, Is.Null);
            Assert.That(sourceExcelRangeE3.Comment, Is.Null);
            sourceExcelRangeC3.AddComment("Testing comment 1", "test1");
            sourceExcelRangeD3.AddComment("Testing comment 2", "test1");
            sourceExcelRangeE3.AddComment("Testing comment 3", "test1");
            Assert.That("test1", Is.EqualTo(sourceExcelRangeC3.Comment.Author));
            Assert.That("Testing comment 1", Is.EqualTo(sourceExcelRangeC3.Comment.Text));
            Assert.That("test1", Is.EqualTo(sourceExcelRangeD3.Comment.Author));
            Assert.That("Testing comment 2", Is.EqualTo(sourceExcelRangeD3.Comment.Text));
            Assert.That("test1", Is.EqualTo(sourceExcelRangeE3.Comment.Author));
            Assert.That("Testing comment 3", Is.EqualTo(sourceExcelRangeE3.Comment.Text));
            // Copy the full row to capture each cell at once.
            Assert.Equals(null, ws1.Cells[5, 3].Comment);
            Assert.Equals(null, ws1.Cells[5, 4].Comment);
            Assert.Equals(null, ws1.Cells[5, 5].Comment);
            ws1.Cells["3:3"].Copy(ws1.Cells["5:5"]);
            // Assert the original comments are intact.
            Assert.That("test1", Is.EqualTo(sourceExcelRangeC3.Comment.Author));
            Assert.That("Testing comment 1", Is.EqualTo(sourceExcelRangeC3.Comment.Text));
            Assert.That("test1", Is.EqualTo(sourceExcelRangeD3.Comment.Author));
            Assert.That("Testing comment 2", Is.EqualTo(sourceExcelRangeD3.Comment.Text));
            Assert.That("test1", Is.EqualTo(sourceExcelRangeE3.Comment.Author));
            Assert.That("Testing comment 3", Is.EqualTo(sourceExcelRangeE3.Comment.Text));
            // Assert the comments were copied.
            var destinationExcelRangeC5 = ws1.Cells[5, 3];
            var destinationExcelRangeD5 = ws1.Cells[5, 4];
            var destinationExcelRangeE5 = ws1.Cells[5, 5];
            Assert.That("test1", Is.EqualTo(destinationExcelRangeC5.Comment.Author));
            Assert.That("Testing comment 1", Is.EqualTo(destinationExcelRangeC5.Comment.Text));
            Assert.That("test1", Is.EqualTo(destinationExcelRangeD5.Comment.Author));
            Assert.That("Testing comment 2", Is.EqualTo(destinationExcelRangeD5.Comment.Text));
            Assert.That("test1", Is.EqualTo(destinationExcelRangeE5.Comment.Author));
            Assert.That("Testing comment 3", Is.EqualTo(destinationExcelRangeE5.Comment.Text));
        }

        [Test]
        public void SettingAddressHandlesMultiAddresses()
        {
            using (ExcelPackage package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                var name = package.Workbook.Names.Add("Test", worksheet.Cells[3, 3]);
                name.Address = "Sheet1!C3";
                name.Address = "Sheet1!D3";
                Assert.That(name.Addresses, Is.Null);
                name.Address = "C3:D3,E3:F3";
                Assert.That(name.Addresses, Is.Not.Null);
                name.Address = "Sheet1!C3";
                Assert.That(name.Addresses, Is.Null);
            }
        }
    }
}
