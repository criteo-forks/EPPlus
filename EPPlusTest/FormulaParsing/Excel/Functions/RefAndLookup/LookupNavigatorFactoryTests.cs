using System;
using System.IO;
using EPPlusTest.FormulaParsing.TestHelpers;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.FormulaParsing.Excel.Functions.RefAndLookup
{
    [TestFixture]
    public class LookupNavigatorFactoryTests
    {
        private ExcelPackage _excelPackage;
        private ParsingContext _context;

        [SetUp]
        public void Initialize()
        {
            _excelPackage = new ExcelPackage(new MemoryStream());
            _excelPackage.Workbook.Worksheets.Add("Test");
            _context = ParsingContext.Create();
            _context.ExcelDataProvider = new EpplusExcelDataProvider(_excelPackage);
            _context.Scopes.NewScope(RangeAddress.Empty);
        }

        [TearDown]
        public void Cleanup()
        {
            _excelPackage.Dispose();
        }

        [Test]
        public void Should_Return_ExcelLookupNavigator_When_Range_Is_Set()
        {
            var args = new LookupArguments(FunctionsHelper.CreateArgs(8, "A:B", 1), ParsingContext.Create());
            var navigator = LookupNavigatorFactory.Create(LookupDirection.Horizontal, args, _context);
            Assert.That(navigator, Is.InstanceOf<ExcelLookupNavigator>());
        }

        [Test]
        public void Should_Return_ArrayLookupNavigator_When_Array_Is_Supplied()
        {
            var args = new LookupArguments(FunctionsHelper.CreateArgs(8, FunctionsHelper.CreateArgs(1,2), 1), ParsingContext.Create());
            var navigator = LookupNavigatorFactory.Create(LookupDirection.Horizontal, args, _context);
            Assert.That(navigator, Is.InstanceOf<ArrayLookupNavigator>());
        }
    }
}
