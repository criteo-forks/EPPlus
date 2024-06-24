using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestFixture]
    public class TimeFormulaTests : ValidationTestBase
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
        }

        [Test]
        public void TimeFormula_ValueIsSetFromConstructorValidateHour()
        {
            // Arrange
            var time = new ExcelTime(0.675M);
            LoadXmlTestData("A1", "time", "0.675");

            // Act
            var formula = new ExcelDataValidationTime(_sheet, "A1", ExcelDataValidationType.Time, _dataValidationNode, _namespaceManager);
            
            // Assert
            Assert.That(time.Hour, Is.EqualTo(formula.Formula.Value.Hour));
        }

        [Test]
        public void TimeFormula_ValueIsSetFromConstructorValidateMinute()
        {
            // Arrange
            var time = new ExcelTime(0.395M);
            LoadXmlTestData("A1", "time", "0.395");

            // Act
            var formula = new ExcelDataValidationTime(_sheet, "A1", ExcelDataValidationType.Time, _dataValidationNode, _namespaceManager);

            // Assert
            Assert.That(time.Minute, Is.EqualTo(formula.Formula.Value.Minute));
        }

        [Test]
        public void TimeFormula_ValueIsSetFromConstructorValidateSecond()
        {
            // Arrange
            var time = new ExcelTime(0.812M);
            LoadXmlTestData("A1", "time", "0.812");

            // Act
            var formula = new ExcelDataValidationTime(_sheet, "A1", ExcelDataValidationType.Time, _dataValidationNode, _namespaceManager);

            // Assert
            Assert.That(time.Second.Value, Is.EqualTo(formula.Formula.Value.Second.Value));
        }
    }
}
