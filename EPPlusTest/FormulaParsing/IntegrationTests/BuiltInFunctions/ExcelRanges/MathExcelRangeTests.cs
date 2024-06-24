using System;
using NUnit.Framework;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions.ExcelRanges
{
    [TestFixture]
    public class MathExcelRangeTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;

        [SetUp]
        public void Initialize()
        {
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
        }

        [Test]
        public void AbsShouldReturn3()
        {
            _worksheet.Cells["A4"].Formula = "ABS(A2)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(3d, Is.EqualTo(result));
        }

        [Test]
        public void CountShouldReturn3()
        {
            _worksheet.Cells["A4"].Formula = "COUNT(A1:A3)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(3d, Is.EqualTo(result));
        }

        [Test]
        public void CountShouldReturn2IfACellValueIsNull()
        {
            _worksheet.Cells["A2"].Value = null;
            _worksheet.Cells["A4"].Formula = "COUNT(A1:A3)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(2d, Is.EqualTo(result));
        }

        [Test]
        public void CountAShouldReturn3()
        {
            _worksheet.Cells["A4"].Formula = "COUNTA(A1:A3)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(3d, Is.EqualTo(result));
        }

        [Test]
        public void CountIfShouldReturnCorrectResult()
        {
            _worksheet.Cells["A4"].Formula = "COUNTIF(A1:A3, \">2\")";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(2d, Is.EqualTo(result));
        }

        [Test]
        public void MaxShouldReturn6()
        {
            _worksheet.Cells["A4"].Formula = "Max(A1:A3)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(6d, Is.EqualTo(result));
        }

        [Test]
        public void MinShouldReturn1()
        {
            _worksheet.Cells["A4"].Formula = "Min(A1:A3)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(1d, Is.EqualTo(result));
        }

        [Test]
        public void AverageShouldReturn3Point333333()
        {
            _worksheet.Cells["A4"].Formula = "Average(A1:A3)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(3d + (1d/3d), Is.EqualTo(result));
        }

        [Test]
        public void AverageIfShouldHandleSingleRangeNumericExpressionMatch()
        {
            _worksheet.Cells["A4"].Value = "B";
            _worksheet.Cells["A5"].Value = 3;
            _worksheet.Cells["A6"].Formula = "AverageIf(A1:A5,\">1\")";
            _worksheet.Calculate();
            Assert.That(4d, Is.EqualTo(_worksheet.Cells["A6"].Value));
        }

        [Test]
        public void AverageIfShouldHandleSingleRangeStringMatch()
        {
            _worksheet.Cells["A4"].Value = "ABC";
            _worksheet.Cells["A5"].Value = "3";
            _worksheet.Cells["A6"].Formula = "AverageIf(A1:A5,\">1\")";
            _worksheet.Calculate();
            Assert.That(4.5d, Is.EqualTo(_worksheet.Cells["A6"].Value));
        }

        [Test]
        public void AverageIfShouldHandleLookupRangeStringMatch()
        {
            _worksheet.Cells["A1"].Value = "abc";
            _worksheet.Cells["A2"].Value = "abc";
            _worksheet.Cells["A3"].Value = "def";
            _worksheet.Cells["A4"].Value = "def";
            _worksheet.Cells["A5"].Value = "abd";

            _worksheet.Cells["B1"].Value = 1;
            _worksheet.Cells["B2"].Value = 3;
            _worksheet.Cells["B3"].Value = 5;
            _worksheet.Cells["B4"].Value = 6;
            _worksheet.Cells["B5"].Value = 7;

            _worksheet.Cells["A6"].Formula = "AverageIf(A1:A5,\"abc\",B1:B5)";
            _worksheet.Calculate();
            Assert.That(2d, Is.EqualTo(_worksheet.Cells["A6"].Value));
        }

        [Test]
        public void AverageIfShouldHandleLookupRangeStringNumericMatch()
        {
            _worksheet.Cells["A1"].Value = 1;
            _worksheet.Cells["A2"].Value = 3;
            _worksheet.Cells["A3"].Value = 3;
            _worksheet.Cells["A4"].Value = 5;
            _worksheet.Cells["A5"].Value = 2;

            _worksheet.Cells["B1"].Value = 3;
            _worksheet.Cells["B2"].Value = 3;
            _worksheet.Cells["B3"].Value = 2;
            _worksheet.Cells["B4"].Value = 1;
            _worksheet.Cells["B5"].Value = 8;

            _worksheet.Cells["A6"].Formula = "AverageIf(A1:A5,\">2\",B1:B5)";
            _worksheet.Calculate();
            Assert.That(2d, Is.EqualTo(_worksheet.Cells["A6"].Value));
        }

        [Test]
        public void AverageIfShouldHandleLookupRangeStringWildCardMatch()
        {
            _worksheet.Cells["A1"].Value = "abc";
            _worksheet.Cells["A2"].Value = "abc";
            _worksheet.Cells["A3"].Value = "def";
            _worksheet.Cells["A4"].Value = "def";
            _worksheet.Cells["A5"].Value = "abd";

            _worksheet.Cells["B1"].Value = 1;
            _worksheet.Cells["B2"].Value = 3;
            _worksheet.Cells["B3"].Value = 5;
            _worksheet.Cells["B4"].Value = 6;
            _worksheet.Cells["B5"].Value = 8;

            _worksheet.Cells["A6"].Formula = "AverageIf(A1:A5, \"ab*\",B1:B5)";
            _worksheet.Calculate();
            Assert.That(4d, Is.EqualTo(_worksheet.Cells["A6"].Value));
        }

        [Test]
        public void SumProductWithRange()
        {
            _worksheet.Cells["A1"].Value = 1;
            _worksheet.Cells["A2"].Value = 2;
            _worksheet.Cells["A3"].Value = 3;
            _worksheet.Cells["B1"].Value = 5;
            _worksheet.Cells["B2"].Value = 6;
            _worksheet.Cells["B3"].Value = 4;
            _worksheet.Cells["A4"].Formula = "SUMPRODUCT(A1:A3,B1:B3)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(29d, Is.EqualTo(result));
        }

        [Test]
        public void SumProductWithRangeAndValues()
        {
            _worksheet.Cells["A1"].Value = 1;
            _worksheet.Cells["A2"].Value = 2;
            _worksheet.Cells["A3"].Value = 3;
            _worksheet.Cells["B1"].Value = 5;
            _worksheet.Cells["B2"].Value = 6;
            _worksheet.Cells["B3"].Value = 4;
            _worksheet.Cells["A4"].Formula = "SUMPRODUCT(A1:A3,B1:B3,{2,4,1})";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(70d, Is.EqualTo(result));
        }

        [Test]
        public void SignShouldReturn1WhenRefIsPositive()
        {
            _worksheet.Cells["A4"].Formula = "SIGN(A1)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(1d, Is.EqualTo(result));
        }

        [Test]
        public void SubTotalShouldNotIncludeHiddenRow()
        {
            _worksheet.Row(2).Hidden = true;
            _worksheet.Cells["A4"].Formula = "SUBTOTAL(109,A1:A3)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(7d, Is.EqualTo(result));
        }

        [Test]
        public void SumProductShouldWorkWithSingleCellArray()
        {
            _worksheet.Cells["A1"].Value = 1;
            _worksheet.Cells["A2"].Value = 2;
            _worksheet.Cells["A4"].Formula = "SUMPRODUCT(A1:A1, A2:A2)";
            _worksheet.Calculate();
            var result = _worksheet.Cells["A4"].Value;
            Assert.That(2d, Is.EqualTo(result));
        }

        [Test]
        public void ShouldIgnoreNullValues()
        {
            _worksheet.Cells["B3"].Formula = "C4 + D4";
            _worksheet.Calculate();
            var result = _worksheet.Cells["B3"].Value;
            Assert.That(0d, Is.EqualTo(result));
        }
    }
}
