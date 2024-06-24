using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Logging;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestFixture]
    public class TextFunctionsTests
    {
        [Test]
        public void HyperlinkShouldHandleReference()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "HYPERLINK(B1)";
                sheet.Cells["B1"].Value = "http://epplus.codeplex.com";
                sheet.Calculate();
                Assert.That("http://epplus.codeplex.com", Is.EqualTo(sheet.Cells["A1"].Value));
            }
        }

        [Test]
        public void HyperlinkShouldHandleReference2()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "HYPERLINK(B1, B2)";
                sheet.Cells["B1"].Value = "http://epplus.codeplex.com";
                sheet.Cells["B2"].Value = "Epplus";
                sheet.Calculate();
                Assert.That("Epplus", Is.EqualTo(sheet.Cells["A1"].Value));
            }
        }

        [Test]
        public void HyperlinkShouldHandleText()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "HYPERLINK(\"testing\")";
                sheet.Calculate();
                Assert.That("testing", Is.EqualTo(sheet.Cells["A1"].Value));
            }
        }

        [Test]
        public void CharShouldReturnCharValOfNumber()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Char(A2)";
                sheet.Cells["A2"].Value = 55;
                sheet.Calculate();
                Assert.That("7", Is.EqualTo(sheet.Cells["A1"].Value));
            }
        }

        [Test]
        public void FixedShouldHaveCorrectDefaultValues()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Fixed(A2)";
                sheet.Cells["A2"].Value = 1234.5678;
                sheet.Calculate();
                Assert.That(1234.5678.ToString("N2"), Is.EqualTo(sheet.Cells["A1"].Value));
            }
        }

        [Test]
        public void FixedShouldSetCorrectNumberOfDecimals()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Fixed(A2,4)";
                sheet.Cells["A2"].Value = 1234.56789;
                sheet.Calculate();
                Assert.That(1234.56789.ToString("N4"), Is.EqualTo(sheet.Cells["A1"].Value));
            }
        }

        [Test]
        public void FixedShouldSetNoCommas()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Fixed(A2,4,true)";
                sheet.Cells["A2"].Value = 1234.56789;
                sheet.Calculate();
                Assert.That(1234.56789.ToString("F4"), Is.EqualTo(sheet.Cells["A1"].Value));
            }
        }

        [Test]
        public void FixedShouldHandleNegativeDecimals()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Fixed(A2,-1,true)";
                sheet.Cells["A2"].Value = 1234.56789;
                sheet.Calculate();
                Assert.That(1230.ToString("F0"), Is.EqualTo(sheet.Cells["A1"].Value));
            }
        }

        [Test]
        public void ConcatenateShouldHandleRange()
        {
            using (var pck = new ExcelPackage(new MemoryStream()))
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "Concatenate(1,A2)";
                sheet.Cells["A2"].Value = "hello";
                sheet.Calculate();
                Assert.That("1hello", Is.EqualTo(sheet.Cells["A1"].Value));
            }
        }

        [Test] [Explicit]
        public void Logtest1()
        {
            var sw = new Stopwatch();
            sw.Start();
            using (var pck = new ExcelPackage(new FileInfo(@"c:\temp\denis.xlsx")))
            {
                var logger = LoggerFactory.CreateTextFileLogger(new FileInfo(@"c:\temp\log1.txt"));
                pck.Workbook.FormulaParser.Configure(x => x.AttachLogger(logger));
                pck.Workbook.Calculate();
                //
            }
            sw.Stop();
            var elapsed = sw.Elapsed;
            Console.WriteLine(string.Format("{0} seconds", elapsed.TotalSeconds));
        }
    }
}
