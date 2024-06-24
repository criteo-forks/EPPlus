using System;
using System.IO;
using EPPlusTest.FormulaParsing.TestHelpers;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.Exceptions;
using Index = OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup.Index;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestFixture]
    public class IndexTests
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
        public void Index_Should_Return_Value_By_Index()
        {
            var func = new Index();
            var result = func.Execute(
                FunctionsHelper.CreateArgs(
                    FunctionsHelper.CreateArgs(1, 2, 5),
                    3
                    ),_parsingContext);
            Assert.That(5, Is.EqualTo(result.Result));
        }

        [Test]
        public void Index_Should_Handle_SingleRange()
        {
            _worksheet.Cells["A1"].Value = 1d;
            _worksheet.Cells["A2"].Value = 3d;
            _worksheet.Cells["A3"].Value = 5d;

            _worksheet.Cells["A4"].Formula = "INDEX(A1:A3;3)";

            _worksheet.Calculate();

            Assert.That(5d, Is.EqualTo(_worksheet.Cells["A4"].Value));
        }
    }
}
