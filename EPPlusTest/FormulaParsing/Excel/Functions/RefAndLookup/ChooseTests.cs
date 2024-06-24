using System;
using System.IO;
using EPPlusTest.FormulaParsing.TestHelpers;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestFixture]
    public class ChooseTests
    {
        private ParsingContext _parsingContext;
        private ExcelPackage _package;
        private ExcelWorksheet _worksheet;

        [SetUp]
        public void Initialize()
        {
            _parsingContext = ParsingContext.Create();
            _package = new ExcelPackage(new MemoryStream());
            _worksheet = _package.Workbook.Worksheets.Add("test");
        }

        [TearDown]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [Test]
        public void ChooseSingleValue()
        {
            fillChooseOptions();
            _worksheet.Cells["B1"].Formula = "CHOOSE(4, A1, A2, A3, A4, A5)";
            _worksheet.Calculate();

            Assert.That("5", Is.EqualTo(_worksheet.Cells["B1"].Value));
        }

        [Test]
        public void ChooseSingleFormula()
        {
            fillChooseOptions();
            _worksheet.Cells["B1"].Formula = "CHOOSE(6, A1, A2, A3, A4, A5, A6)";
            _worksheet.Calculate();

            Assert.That("12", Is.EqualTo(_worksheet.Cells["B1"].Value));
        }

        [Test]
        public void ChooseMultipleValues()
        {
            fillChooseOptions();
            _worksheet.Cells["B1"].Formula = "SUM(CHOOSE({1,3,4}, A1, A2, A3, A4, A5))";
            _worksheet.Calculate();

            Assert.That(9D, Is.EqualTo(_worksheet.Cells["B1"].Value));
        }

        [Test]
        public void ChooseValueAndFormula()
        {
            fillChooseOptions();
            _worksheet.Cells["B1"].Formula = "SUM(CHOOSE({2,6}, A1, A2, A3, A4, A5, A6))";
            _worksheet.Calculate();

            Assert.That(14D, Is.EqualTo(_worksheet.Cells["B1"].Value));
        }

        private void fillChooseOptions()
        {
            _worksheet.Cells["A1"].Value = 1d;
            _worksheet.Cells["A2"].Value = 2d;
            _worksheet.Cells["A3"].Value = 3d;
            _worksheet.Cells["A4"].Value = 5d;
            _worksheet.Cells["A5"].Value = 7d;
            _worksheet.Cells["A6"].Formula = "A4 + A5";
        }
    }
}
