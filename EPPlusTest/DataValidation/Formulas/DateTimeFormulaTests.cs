using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.DataValidation;
using OfficeOpenXml;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestFixture]
    public class DateTimeFormulaTests : ValidationTestBase
    {
        [SetUp]
        public void Setup()
        {
            SetupTestData();
        }

        [TearDown]
        public void Cleanup()
        {
            CleanupTestData();
            _dataValidationNode = null;
        }

        [Test]
        public void DateTimeFormula_FormulaValueIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            var date = DateTime.Parse("2011-01-08");
            var dateAsString = date.ToOADate().ToString(_cultureInfo);
            LoadXmlTestData("A1", "decimal", dateAsString);
            // Act
            var validation = new ExcelDataValidationDateTime(_sheet, "A1", ExcelDataValidationType.Decimal, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.That(date, Is.EqualTo(validation.Formula.Value));
        }

        [Test]
        public void DateTimeFormula_FormulasFormulaIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            var date = DateTime.Parse("2011-01-08");
            LoadXmlTestData("A1", "decimal", "A1");

            // Act
            var validation = new ExcelDataValidationDateTime(_sheet, "A1", ExcelDataValidationType.Decimal, _dataValidationNode, _namespaceManager);

            // Assert
            Assert.That("A1", Is.EqualTo(validation.Formula.ExcelFormula));
        }
    }
}
