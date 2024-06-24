using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using NUnit.Framework;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions.ExcelRanges
{
    [TestFixture]
    public class TextExcelRangeTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;
        private CultureInfo _currentCulture;

        [SetUp]
        public void Initialize()
        {
            _currentCulture = CultureInfo.CurrentCulture;
            _package = new ExcelPackage();
            _worksheet = _package.Workbook.Worksheets.Add("Test");

            _worksheet.Cells["A1"].Value = 1;
            _worksheet.Cells["A2"].Value = 3;
            _worksheet.Cells["A3"].Value = 6;
        }

        [TearDown]
        public void Cleanup()
        {
            _package.Dispose();
            Thread.CurrentThread.CurrentCulture = _currentCulture;
        }

        [Test]
        public void ExactShouldReturnTrueWhenEqualValues()
        {
            _worksheet.Cells["A2"].Value = 1d;
            _worksheet.Cells["A4"].Formula = "EXACT(A1,A2)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That((bool)result);
        }

        [Test]
        public void FindShouldReturnIndexCaseSensitive()
        {
            _worksheet.Cells["A1"].Value = "h";
            _worksheet.Cells["A2"].Value = "Hej hopp";
            _worksheet.Cells["A4"].Formula = "Find(A1,A2)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(5, Is.EqualTo(result));
        }

        [Test]
        public void SearchShouldReturnIndexCaseInSensitive()
        {
            _worksheet.Cells["A1"].Value = "h";
            _worksheet.Cells["A2"].Value = "Hej hopp";
            _worksheet.Cells["A4"].Formula = "Search(A1,A2)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(1, Is.EqualTo(result));
        }

        [Test]
        public void ValueShouldHandleStringWithIntegers()
        {
            _worksheet.Cells["A1"].Value = "12";
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(12d, Is.EqualTo(result));
        }

        [Test]
        public void ValueShouldHandle1000delimiter()
        {
            var delimiter = CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
            var val = $"5{delimiter}000";
            _worksheet.Cells["A1"].Value = val;
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(5000d, Is.EqualTo(result));
        }

        [Test]
        public void ValueShouldHandle1000DelimiterAndDecimal()
        {
            var delimiter = CultureInfo.CurrentCulture.NumberFormat.NumberGroupSeparator;
            var decimalSeparator = CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator;
            var val = $"5{delimiter}000{decimalSeparator}123";
            _worksheet.Cells["A1"].Value = val;
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(5000.123d, Is.EqualTo(result));
        }

        [Test]
        public void ValueShouldHandlePercent()
        {
            var val = $"20%";
            _worksheet.Cells["A1"].Value = val;
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(0.2d, Is.EqualTo(result));
        }

        [Test]
        public void ValueShouldHandleScientificNotation()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            _worksheet.Cells["A1"].Value = "1.2345E-02";
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(0.012345d, Is.EqualTo(result));
        }

        [Test]
        public void ValueShouldHandleDate()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var date = new DateTime(2015, 12, 31);
            _worksheet.Cells["A1"].Value = date.ToString(CultureInfo.CurrentCulture);
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(date.ToOADate(), Is.EqualTo(result));
        }

        [Test]
        public void ValueShouldHandleTime()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var date = new DateTime(2015, 12, 31);
            var date2 = new DateTime(2015, 12, 31, 12, 00, 00);
            var ts = date2.Subtract(date);
            _worksheet.Cells["A1"].Value = ts.ToString();
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(0.5, Is.EqualTo(result));
        }

        [Test]
        public void ValueShouldReturn0IfValueIsNull()
        {

            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A4"].Formula = "Value(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(0d, Is.EqualTo(result));
        }

    }
}
