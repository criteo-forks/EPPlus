using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.DataValidation;

namespace EPPlusTest.DataValidation
{
    [TestFixture]
    public class ValidationCollectionTests : ValidationTestBase
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
        public void ExcelDataValidationCollection_AddDecimal_ShouldThrowWhenAddressIsNullOrEmpty()
        {
            Assert.Throws<ArgumentNullException>(() =>
            {
                // Act
                _sheet.DataValidations.AddDecimalValidation(string.Empty);
            });
            
        }

        [Test]
        public void ExcelDataValidationCollection_AddDecimal_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            Assert.Throws<InvalidOperationException>(() =>
            {
                // Act
                _sheet.DataValidations.AddDecimalValidation("A1");
                _sheet.DataValidations.AddDecimalValidation("A1");
            });
            
        }

        [Test]
        public void ExcelDataValidationCollection_AddInteger_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            Assert.Throws<InvalidOperationException>(() =>
            {
                // Act
                _sheet.DataValidations.AddIntegerValidation("A1");
                _sheet.DataValidations.AddIntegerValidation("A1");
            });
            
        }

        [Test]
        public void ExcelDataValidationCollection_AddTextLength_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            Assert.Throws<InvalidOperationException>(() =>
            {
                // Act
                _sheet.DataValidations.AddTextLengthValidation("A1");
                _sheet.DataValidations.AddTextLengthValidation("A1");
            });
            
        }

        [Test]
        public void ExcelDataValidationCollection_AddDateTime_ShouldThrowWhenNewValidationCollidesWithExisting()
        {
            Assert.Throws<InvalidOperationException>(() =>
            {
                // Act
                _sheet.DataValidations.AddDateTimeValidation("A1");
                _sheet.DataValidations.AddDateTimeValidation("A1");
            });
        }

        [Test]
        public void ExcelDataValidationCollection_Index_ShouldReturnItemAtIndex()
        {
            // Arrange
            _sheet.DataValidations.AddDateTimeValidation("A1");
            _sheet.DataValidations.AddDateTimeValidation("A2");
            _sheet.DataValidations.AddDateTimeValidation("B1");

            // Act
            var result = _sheet.DataValidations[1];

            // Assert
            Assert.That("A2", Is.EqualTo(result.Address.Address));
        }

        [Test]
        public void ExcelDataValidationCollection_FindAll_ShouldReturnValidationInColumnAonly()
        {
            // Arrange
            _sheet.DataValidations.AddDateTimeValidation("A1");
            _sheet.DataValidations.AddDateTimeValidation("A2");
            _sheet.DataValidations.AddDateTimeValidation("B1");

            // Act
            var result = _sheet.DataValidations.FindAll(x => x.Address.Address.StartsWith("A"));

            // Assert
            Assert.That(2, Is.EqualTo(result.Count()));

        }

        [Test]
        public void ExcelDataValidationCollection_Find_ShouldReturnFirstMatchOnly()
        {
            // Arrange
            _sheet.DataValidations.AddDateTimeValidation("A1");
            _sheet.DataValidations.AddDateTimeValidation("A2");

            // Act
            var result = _sheet.DataValidations.Find(x => x.Address.Address.StartsWith("A"));

            // Assert
            Assert.That("A1", Is.EqualTo(result.Address.Address));

        }

        [Test]
        public void ExcelDataValidationCollection_Clear_ShouldBeEmpty()
        {
            // Arrange
            _sheet.DataValidations.AddDateTimeValidation("A1");

            // Act
            _sheet.DataValidations.Clear();

            // Assert
            Assert.That(0, Is.EqualTo(_sheet.DataValidations.Count));

        }

        [Test]
        public void ExcelDataValidationCollection_RemoveAll_ShouldRemoveMatchingEntries()
        {
            // Arrange
            _sheet.DataValidations.AddIntegerValidation("A1");
            _sheet.DataValidations.AddIntegerValidation("A2");
            _sheet.DataValidations.AddIntegerValidation("B1");

            // Act
            _sheet.DataValidations.RemoveAll(x => x.Address.Address.StartsWith("B"));

            // Assert
            Assert.That(2, Is.EqualTo(_sheet.DataValidations.Count));
        }
    }
}
