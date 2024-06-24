using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.DataValidation.Formulas;
using System.Xml;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestFixture]
    public class DecimalFormulaTests : ValidationTestBase
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
        public void DecimalFormula_FormulaValueIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            LoadXmlTestData("A1", "decimal", "1.3");
            // Act
            var validation = new ExcelDataValidationDecimal(_sheet, "A1", ExcelDataValidationType.Decimal, _dataValidationNode, _namespaceManager);
            Assert.That(1.3D, Is.EqualTo(validation.Formula.Value));
        }

        [Test]
        public void DecimalFormula_FormulasFormulaIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            LoadXmlTestData("A1", "decimal", "A1");

            // Act
            var validation = new ExcelDataValidationDecimal(_sheet, "A1", ExcelDataValidationType.Decimal, _dataValidationNode, _namespaceManager);

            // Assert
            Assert.That("A1", Is.EqualTo(validation.Formula.ExcelFormula));
        }
    }
}
