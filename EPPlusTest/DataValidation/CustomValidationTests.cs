using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;

namespace EPPlusTest.DataValidation
{
    [TestFixture]
    public class CustomValidationTests : ValidationTestBase
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
        public void CustomValidation_FormulaIsSet()
        {
            // Act
            var validation = _sheet.DataValidations.AddCustomValidation("A1");

            // Assert
            Assert.That(validation.Formula, Is.Not.Null);
        }

        [Test]
        public void CustomValidation_ShouldThrowExceptionIfFormulaIsTooLong()
        {
            Assert.Throws<InvalidOperationException>(() =>
            {
                // Arrange
                var sb = new StringBuilder();
                for (var x = 0; x < 257; x++) sb.Append("x");

                // Act
                var validation = _sheet.DataValidations.AddCustomValidation("A1");
                validation.Formula.ExcelFormula = sb.ToString();
            });
        }
    }
}
