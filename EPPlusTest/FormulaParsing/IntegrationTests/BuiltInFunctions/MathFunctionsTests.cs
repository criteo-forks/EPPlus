using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestFixture]
    public class MathFunctionsTests : FormulaParserTestBase
    {
        private ExcelPackage _package;

        [SetUp]
        public void Setup()
        {
            _package = new ExcelPackage();
            var excelDataProvider = new EpplusExcelDataProvider(_package);
            _parser = new FormulaParser(excelDataProvider);
        }

        [TearDown]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [Test]
        public void PowerShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Power(3, 3)");
            Assert.That(27d, Is.EqualTo(result));
        }

        [Test]
        public void SqrtShouldReturnCorrectResult()
        {
            var result = _parser.Parse("sqrt(9)");
            Assert.That(3d, Is.EqualTo(result));
        }

        [Test]
        public void PiShouldReturnCorrectResult()
        {
            var expectedValue = (double)Math.Round(Math.PI, 14);
            var result = _parser.Parse("Pi()");
            Assert.That(expectedValue, Is.EqualTo(result));
        }

        [Test]
        public void CeilingShouldReturnCorrectResult()
        {
            var expectedValue = 22.4d;
            var result = _parser.Parse("ceiling(22.35, 0.1)");
            Assert.That(expectedValue, Is.EqualTo(result));
        }

        [Test]
        public void FloorShouldReturnCorrectResult()
        {
            var expectedValue = 22.3d;
            var result = _parser.Parse("Floor(22.35, 0.1)");
            Assert.That(expectedValue, Is.EqualTo(result));
        }

        [Test]
        public void SumShouldReturnCorrectResultWithInts()
        {
            var result = _parser.Parse("sum(1, 2)");
            Assert.That(3d, Is.EqualTo(result));
        }

        [Test]
        public void SumShouldReturnCorrectResultWithDecimals()
        {
            var result = _parser.Parse("sum(1,2.5)");
            Assert.That(3.5d, Is.EqualTo(result));
        }

        [Test]
        public void SumShouldReturnCorrectResultWithEnumerable()
        {
            var result = _parser.Parse("sum({1;2;3;-1}, 2.5)");
            Assert.That(7.5d, Is.EqualTo(result));
        }

        [Test]
        public void SumsqShouldReturnCorrectResultWithEnumerable()
        {
            var result = _parser.Parse("sumsq({2;3})");
            Assert.That(13d, Is.EqualTo(result));
        }

        [Test]
        public void SubtotalShouldNegateExpression()
        {
            var result = _parser.Parse("-subtotal(2;{1;2})");
            Assert.That(-2d, Is.EqualTo(result));
        }

        [Test]
        public void StdevShouldReturnAResult()
        {
            var result = _parser.Parse("stdev(1;2;3;4)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void StdevPShouldReturnAResult()
        {
            var result = _parser.Parse("stdevp(2,3,4)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void ExpShouldReturnAResult()
        {
            var result = _parser.Parse("exp(4)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void MaxShouldReturnAResult()
        {
            var result = _parser.Parse("Max(4, 5)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void MaxaShouldReturnAResult()
        {
            var result = _parser.Parse("Maxa(4, 5)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void MinShouldReturnAResult()
        {
            var result = _parser.Parse("min(4, 5)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void MinaShouldCalculateStringAs0()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B2"].Value = "a";
                sheet.Cells["A5"].Formula = "MINA(A1:B4)";
                sheet.Calculate();
                Assert.That(0d, Is.EqualTo(sheet.Cells["A5"].Value));
            }
        }

        [Test]
        public void AverageShouldReturnAResult()
        {
            var result = _parser.Parse("Average(2, 2, 2)");
            Assert.That(2d, Is.EqualTo(result));
        }

        [Test]
        public void AverageShouldReturnDiv0IfEmptyCell()
        {
            using(var pck = new ExcelPackage())
            {
                var ws = pck.Workbook.Worksheets.Add("test");
                ws.Cells["A2"].Formula = "AVERAGE(A1)";
                ws.Calculate();
                Assert.That("#DIV/0!", Is.EqualTo(ws.Cells["A2"].Value.ToString()));
            }
        }

        [Test]
        public void RoundShouldReturnAResult()
        {
            var result = _parser.Parse("Round(2.2, 0)");
            Assert.That(2d, Is.EqualTo(result));
        }

        [Test]
        public void RounddownShouldReturnAResult()
        {
            var result = _parser.Parse("Rounddown(2.99, 1)");
            Assert.That(2.9d, Is.EqualTo(result));
        }

        [Test]
        public void RoundupShouldReturnAResult()
        {
            var result = _parser.Parse("Roundup(2.99, 1)");
            Assert.That(3d, Is.EqualTo(result));
        }

        [Test]
        public void SqrtPiShouldReturnAResult()
        {
            var result = _parser.Parse("SqrtPi(2.2)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void IntShouldReturnAResult()
        {
            var result = _parser.Parse("Int(2.9)");
            Assert.That(2, Is.EqualTo(result));
        }

        [Test]
        public void RandShouldReturnAResult()
        {
            var result = _parser.Parse("Rand()");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void RandBetweenShouldReturnAResult()
        {
            var result = _parser.Parse("RandBetween(1,2)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void CountShouldReturnAResult()
        {
            var result = _parser.Parse("Count(1,2,2,\"4\")");
            Assert.That(4d, Is.EqualTo(result));
        }

        [Test]
        public void CountAShouldReturnAResult()
        {
            var result = _parser.Parse("CountA(1,2,2,\"\", \"a\")");
            Assert.That(4d, Is.EqualTo(result));
        }

        [Test]
        public void CountIfShouldReturnAResult()
        {
            var result = _parser.Parse("CountIf({1;2;2;\"\"}, \"2\")");
            Assert.That(2d, Is.EqualTo(result));
        }

        [Test]
        public void VarShouldReturnAResult()
        {
            var result = _parser.Parse("Var(1,2,3)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void VarPShouldReturnAResult()
        {
            var result = _parser.Parse("VarP(1,2,3)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void ModShouldReturnAResult()
        {
            var result = _parser.Parse("Mod(5,2)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void SubtotalShouldReturnAResult()
        {
            var result = _parser.Parse("Subtotal(1, 10, 20)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void TruncShouldReturnAResult()
        {
            var result = _parser.Parse("Trunc(1.2345)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void ProductShouldReturnAResult()
        {
            var result = _parser.Parse("Product(1,2,3)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void CosShouldReturnAResult()
        {
            var result = _parser.Parse("Cos(2)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void CoshShouldReturnAResult()
        {
            var result = _parser.Parse("Cosh(2)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void SinShouldReturnAResult()
        {
            var result = _parser.Parse("Sin(2)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void SinhShouldReturnAResult()
        {
            var result = _parser.Parse("Sinh(2)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void TanShouldReturnAResult()
        {
            var result = _parser.Parse("Tan(2)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void AtanShouldReturnAResult()
        {
            var result = _parser.Parse("Atan(2)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void Atan2ShouldReturnAResult()
        {
            var result = _parser.Parse("Atan2(2,1)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void TanhShouldReturnAResult()
        {
            var result = _parser.Parse("Tanh(2)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void LogShouldReturnAResult()
        {
            var result = _parser.Parse("Log(2, 2)");
            Assert.That(1d, Is.EqualTo(result));
        }

        [Test]
        public void Log10ShouldReturnAResult()
        {
            var result = _parser.Parse("Log10(2)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void LnShouldReturnAResult()
        {
            var result = _parser.Parse("Ln(2)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void FactShouldReturnAResult()
        {
            var result = _parser.Parse("Fact(0)");
            Assert.That(1d, Is.EqualTo(result));
        }

        [Test]
        public void QuotientShouldReturnAResult()
        {
            var result = _parser.Parse("Quotient(5;2)");
            Assert.That(2, Is.EqualTo(result));
        }

        [Test]
        public void MedianShouldReturnAResult()
        {
            var result = _parser.Parse("Median(1;2;3)");
            Assert.That(2d, Is.EqualTo(result));
        }

        [Test]
        public void CountBlankShouldCalculateEmptyCells()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B2"].Value = string.Empty;
                sheet.Cells["A5"].Formula = "COUNTBLANK(A1:B4)";
                sheet.Calculate();
                Assert.That(7, Is.EqualTo(sheet.Cells["A5"].Value));
            }
        }

        [Test]
        public void CountBlankShouldCalculateResultOfOffset()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 1;
                sheet.Cells["B2"].Value = string.Empty;
                sheet.Cells["A5"].Formula = "COUNTBLANK(OFFSET(A1, 0, 1))";
                sheet.Calculate();
                Assert.That(1, Is.EqualTo(sheet.Cells["A5"].Value));
            }
        }

        [Test]
        public void DegreesShouldReturnCorrectResult()
        {
            var result = _parser.Parse("DEGREES(0.5)");
            var rounded = Math.Round((double)result, 3);
            Assert.That(28.648, Is.EqualTo(rounded));
        }

        [Test]
        public void AverateIfsShouldCaluclateResult()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["F4"].Value = 1;
                sheet.Cells["F5"].Value = 2;
                sheet.Cells["F6"].Formula = "2 + 2";
                sheet.Cells["F7"].Value = 4;
                sheet.Cells["F8"].Value = 5;

                sheet.Cells["H4"].Value = 3;
                sheet.Cells["H5"].Value = 3;
                sheet.Cells["H6"].Formula = "2 + 2";
                sheet.Cells["H7"].Value = 4;
                sheet.Cells["H8"].Value = 5;

                sheet.Cells["I4"].Value = 2;
                sheet.Cells["I5"].Value = 3;
                sheet.Cells["I6"].Formula = "2 + 2";
                sheet.Cells["I7"].Value = 5;
                sheet.Cells["I8"].Value = 1;

                sheet.Cells["H9"].Formula = "AVERAGEIFS(F4:F8;H4:H8;\">3\";I4:I8;\"<5\")";
                sheet.Calculate();
                Assert.That(4.5d, Is.EqualTo(sheet.Cells["H9"].Value));
            }
        }

        [Test]
        public void AbsShouldHandleEmptyCell()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "ABS(B1)";
                sheet.Calculate();

                Assert.That(0d, Is.EqualTo(sheet.Cells["A1"].Value));
            }
        }
    }
}
