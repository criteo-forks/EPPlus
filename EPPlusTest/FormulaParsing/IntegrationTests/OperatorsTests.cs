using System;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.IntegrationTests
{
    [TestFixture]
    public class OperatorsTests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _ws;
        private readonly ExcelErrorValue DivByZero = ExcelErrorValue.Create(eErrorType.Div0);

        [SetUp]
        public void Initialize()
        {
            _package = new ExcelPackage();
            _ws = _package.Workbook.Worksheets.Add("test");
        }

        [TearDown]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [Test]
        public void DivByZeroShouldReturnError()
        {
            var result = _ws.Calculate("10/0 + 3");
            Assert.That(DivByZero, Is.EqualTo(result));
        }
    }
}
