using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.DataValidation;
using System.Collections;
using NUnit.Framework.Legacy;

namespace EPPlusTest.DataValidation.Formulas
{
    [TestFixture]
    public class ListFormulaTests : ValidationTestBase
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
        public void ListFormula_FormulaValueIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            LoadXmlTestData("A1", "list", "\"1,2\"");
            // Act
            var validation = new ExcelDataValidationList(_sheet, "A1", ExcelDataValidationType.List, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.That(2, Is.EqualTo(validation.Formula.Values.Count));
        }

        [Test]
        public void ListFormula_FormulaValueIsSetFromXmlNodeInConstructorOrderIsCorrect()
        {
            // Arrange
            LoadXmlTestData("A1", "list", "\"1,2\"");
            // Act
            var validation = new ExcelDataValidationList(_sheet, "A1", ExcelDataValidationType.List, _dataValidationNode, _namespaceManager);
            // Assert
            CollectionAssert.AreEquivalent(new List<string>{ "1", "2"}, (ICollection)validation.Formula.Values);
        }

        [Test]
        public void ListFormula_FormulasExcelFormulaIsSetFromXmlNodeInConstructor()
        {
            // Arrange
            LoadXmlTestData("A1", "list", "A1");
            // Act
            var validation = new ExcelDataValidationList(_sheet, "A1", ExcelDataValidationType.List, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.That("A1", Is.EqualTo(validation.Formula.ExcelFormula));
        }
    }
}
