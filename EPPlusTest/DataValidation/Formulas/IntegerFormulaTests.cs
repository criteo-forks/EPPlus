using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestFixture]
    public class IntegerFormulaTests : ValidationTestBase
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
        public void IntegerFormula_FormulaValueIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            LoadXmlTestData("A1", "decimal", "1");
            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
            Assert.That(1, Is.EqualTo(validation.Formula.Value));
        }

        [Test]
        public void IntegerFormula_FormulasFormulaIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            LoadXmlTestData("A1", "decimal", "A1");

            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);

            // Assert
            Assert.That("A1", Is.EqualTo(validation.Formula.ExcelFormula));
        }
    }
}
