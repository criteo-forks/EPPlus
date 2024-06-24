using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml.DataValidation.Contracts;

namespace EPPlusTest.DataValidation
{
    [TestFixture]
    public class DecimaDataValidationTests : ValidationTestBase
    {
        private IExcelDataValidationDecimal _validation;

        [SetUp]
        public void Setup()
        {
            SetupTestData();
            _validation = _package.Workbook.Worksheets[1].DataValidations.AddDecimalValidation("A1");
        }

        [TearDown]
        public void Cleanup()
        {
            CleanupTestData();
            _validation = null;
        }

        [Test]
        public void DecimalDataValidation_Formula1IsSet()
        {
            Assert.That(_validation.Formula, Is.Not.Null);
        }

        [Test]
        public void DecimalDataValidation_Formula2IsSet()
        {
            Assert.That(_validation.Formula2, Is.Not.Null);
        }
    }
}
