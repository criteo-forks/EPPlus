using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;
using System.IO;

namespace EPPlusTest.DataValidation
{
    [TestFixture]
    public class DataValidationTests : ValidationTestBase
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
        public void DataValidations_ShouldSetOperatorFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "greaterThanOrEqual", "1");
            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.That(ExcelDataValidationOperator.greaterThanOrEqual, Is.EqualTo(validation.Operator));
       }

        [Test]
        public void DataValidations_ShouldThrowIfOperatorIsEqualAndFormula1IsEmpty()
        {
            Assert.Throws<InvalidOperationException>(() =>
            {
                var validations = _sheet.DataValidations.AddIntegerValidation("A1");
                validations.Operator = ExcelDataValidationOperator.equal;
                validations.Validate();
            });
        }

        [Test]
        public void DataValidations_ShouldSetShowErrorMessageFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", true, false);
            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.That(validation.ShowErrorMessage ?? false);
        }

        [Test]
        public void DataValidations_ShouldSetShowInputMessageFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", false, true);
            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.That(validation.ShowInputMessage ?? false);
        }

        [Test]
        public void DataValidations_ShouldSetPromptFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.That("Prompt", Is.EqualTo(validation.Prompt));
        }

        [Test]
        public void DataValidations_ShouldSetPromptTitleFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.That("PromptTitle", Is.EqualTo(validation.PromptTitle));
        }

        [Test]
        public void DataValidations_ShouldSetErrorFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.That("Error", Is.EqualTo(validation.Error));
        }

        [Test]
        public void DataValidations_ShouldSetErrorTitleFromExistingXml()
        {
            // Arrange
            LoadXmlTestData("A1", "whole", "1", "Prompt", "PromptTitle", "Error", "ErrorTitle");
            // Act
            var validation = new ExcelDataValidationInt(_sheet, "A1", ExcelDataValidationType.Whole, _dataValidationNode, _namespaceManager);
            // Assert
            Assert.That("ErrorTitle", Is.EqualTo(validation.ErrorTitle));
        }

        [Test]
        public void DataValidations_ShouldThrowIfOperatorIsBetweenAndFormula2IsEmpty()
        {
            Assert.Throws<InvalidOperationException>(() =>
            {
                var validation = _sheet.DataValidations.AddIntegerValidation("A1");
                validation.Formula.Value = 1;
                validation.Operator = ExcelDataValidationOperator.between;
                validation.Validate();
            });
        }

        [Test]
        public void DataValidations_ShouldAcceptOneItemOnly()
        {
            var validation = _sheet.DataValidations.AddListValidation("A1");
            validation.Formula.Values.Add("1");
            validation.Validate();
        }

        [Test]
        public void ExcelDataValidation_ShouldReplaceLastPartInWholeColumnRangeWithMaxNumberOfRowsOneColumn()
        {
            // Act
            var validation = _sheet.DataValidations.AddIntegerValidation("A:A");

            // Assert
            Assert.That("A1:A" + ExcelPackage.MaxRows.ToString(), Is.EqualTo(validation.Address.Address));
        }

        [Test]
        public void ExcelDataValidation_ShouldReplaceLastPartInWholeColumnRangeWithMaxNumberOfRowsDifferentColumns()
        {
            // Act
            var validation = _sheet.DataValidations.AddIntegerValidation("A:B");

            // Assert
            Assert.That(string.Format("A1:B{0}", ExcelPackage.MaxRows), Is.EqualTo(validation.Address.Address));
        }

    }
}
