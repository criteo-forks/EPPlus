using EPPlusTest.FormulaParsing.TestHelpers;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using static OfficeOpenXml.FormulaParsing.ExcelDataProvider;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
    [TestFixture]
    public class AverageIfTests
    {
        private ExcelPackage _package;
        private EpplusExcelDataProvider _provider;
        private ParsingContext _parsingContext;
        private ExcelWorksheet _worksheet;

        [SetUp]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _provider = new EpplusExcelDataProvider(_package);
            _parsingContext = ParsingContext.Create();
            _parsingContext.Scopes.NewScope(RangeAddress.Empty);
            _worksheet = _package.Workbook.Worksheets.Add("testsheet");
        }

        [TearDown]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [Test]
        public void AverageIfNumeric()
        {
            _worksheet.Cells["A1"].Value = 1d;
            _worksheet.Cells["A2"].Value = 2d;
            _worksheet.Cells["A3"].Value = 3d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new AverageIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">1", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.That(4d, Is.EqualTo(result.Result));
        }

        [Test]
        public void AverageIfNonNumeric()
        {
            _worksheet.Cells["A1"].Value = "Monday";
            _worksheet.Cells["A2"].Value = "Tuesday";
            _worksheet.Cells["A3"].Value = "Thursday";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new AverageIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "T*day", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.That(4d, Is.EqualTo(result.Result));
        }

        [Test]
        public void AverageIfNumericExpression()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = 1d;
            _worksheet.Cells["A3"].Value = "Not Empty";
            var func = new AverageIf();
            IRangeInfo range = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            var args = FunctionsHelper.CreateArgs(range, 1d);
            var result = func.Execute(args, _parsingContext);
            Assert.That(1d, Is.EqualTo(result.Result));
        }

        [Test]
        public void AverageIfEqualToEmptyString()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new AverageIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.That(1d, Is.EqualTo(result.Result));
        }

        [Test]
        public void AverageIfNotEqualToNull()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new AverageIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<>", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.That(4d, Is.EqualTo(result.Result));
        }

        [Test]
        public void AverageIfEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 0d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new AverageIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.That(5d, Is.EqualTo(result.Result));
        }

        [Test]
        public void AverageIfNotEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 0d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new AverageIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<>0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.That(2d, Is.EqualTo(result.Result));
        }

        [Test]
        public void AverageIfGreaterThanZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 1d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new AverageIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.That(5d, Is.EqualTo(result.Result));
        }

        [Test]
        public void AverageIfGreaterThanOrEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = 1d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new AverageIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">=0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.That(5d, Is.EqualTo(result.Result));
        }

        [Test]
        public void AverageIfLessThanZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = -1d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new AverageIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.That(5d, Is.EqualTo(result.Result));
        }

        [Test]
        public void AverageIfLessThanOrEqualToZero()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = -1d;
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new AverageIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<=0", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.That(5d, Is.EqualTo(result.Result));
        }

        [Test]
        public void AverageIfLessThanCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new AverageIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<a", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.That(3d, Is.EqualTo(result.Result));
        }

        [Test]
        public void AverageIfLessThanOrEqualToCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new AverageIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, "<=a", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.That(3d, Is.EqualTo(result.Result));
        }

        [Test]
        public void AverageIfGreaterThanCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new AverageIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">a", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.That(5d, Is.EqualTo(result.Result));
        }

        [Test]
        public void AverageIfGreaterThanOrEqualToCharacter()
        {
            _worksheet.Cells["A1"].Value = null;
            _worksheet.Cells["A2"].Value = string.Empty;
            _worksheet.Cells["A3"].Value = "Not Empty";
            _worksheet.Cells["B1"].Value = 1d;
            _worksheet.Cells["B2"].Value = 3d;
            _worksheet.Cells["B3"].Value = 5d;
            var func = new AverageIf();
            IRangeInfo range1 = _provider.GetRange(_worksheet.Name, 1, 1, 3, 1);
            IRangeInfo range2 = _provider.GetRange(_worksheet.Name, 1, 2, 3, 2);
            var args = FunctionsHelper.CreateArgs(range1, ">=a", range2);
            var result = func.Execute(args, _parsingContext);
            Assert.That(5d, Is.EqualTo(result.Result));
        }
    }
}
