using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestFixture]
    public class CustomFormulaTests : ValidationTestBase
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
        public void CustomFormula_FormulasFormulaIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            LoadXmlTestData("A1", "decimal", "A1");

            // Act
            var validation = new ExcelDataValidationCustom(_sheet, "A1", ExcelDataValidationType.Custom, _dataValidationNode, _namespaceManager);

            // Assert
            Assert.That("A1", Is.EqualTo(validation.Formula.ExcelFormula));
        }
    }
}
