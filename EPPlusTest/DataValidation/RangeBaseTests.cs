using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;

namespace EPPlusTest.DataValidation
{
    [TestFixture]
    public class RangeBaseTests : ValidationTestBase
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
        public void RangeBase_AddIntegerValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddIntegerDataValidation();

            // Assert
            Assert.That(1, Is.EqualTo(_sheet.DataValidations.Count));
        }

        [Test]
        public void RangeBase_AddIntegerValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddIntegerDataValidation();

            // Assert
            Assert.That("A1:A2", Is.EqualTo(_sheet.DataValidations[0].Address.Address));
        }

        [Test]
        public void RangeBase_AddDecimalValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddDecimalDataValidation();

            // Assert
            Assert.That(1, Is.EqualTo(_sheet.DataValidations.Count));
        }

        [Test]
        public void RangeBase_AddDecimalValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddDecimalDataValidation();

            // Assert
            Assert.That("A1:A2", Is.EqualTo(_sheet.DataValidations[0].Address.Address));
        }

        [Test]
        public void RangeBase_AddTextLengthValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddTextLengthDataValidation();

            // Assert
            Assert.That(1, Is.EqualTo(_sheet.DataValidations.Count));
        }

        [Test]
        public void RangeBase_AddTextLengthValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddTextLengthDataValidation();

            // Assert
            Assert.That("A1:A2", Is.EqualTo(_sheet.DataValidations[0].Address.Address));
        }

        [Test]
        public void RangeBase_AddDateTimeValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddDateTimeDataValidation();

            // Assert
            Assert.That(1, Is.EqualTo(_sheet.DataValidations.Count));
        }

        [Test]
        public void RangeBase_AddDateTimeValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddDateTimeDataValidation();

            // Assert
            Assert.That("A1:A2", Is.EqualTo(_sheet.DataValidations[0].Address.Address));
        }

        [Test]
        public void RangeBase_AddListValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddListDataValidation();

            // Assert
            Assert.That(1, Is.EqualTo(_sheet.DataValidations.Count));
        }

        [Test]
        public void RangeBase_AddListValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddListDataValidation();

            // Assert
            Assert.That("A1:A2", Is.EqualTo(_sheet.DataValidations[0].Address.Address));
        }

        [Test]
        public void RangeBase_AdTimeValidation_ValidationIsAdded()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddTimeDataValidation();

            // Assert
            Assert.That(1, Is.EqualTo(_sheet.DataValidations.Count));
        }

        [Test]
        public void RangeBase_AddTimeValidation_AddressIsCorrect()
        {
            // Act
            _sheet.Cells["A1:A2"].DataValidation.AddTimeDataValidation();

            // Assert
            Assert.That("A1:A2", Is.EqualTo(_sheet.DataValidations[0].Address.Address));
        }
    }
}
