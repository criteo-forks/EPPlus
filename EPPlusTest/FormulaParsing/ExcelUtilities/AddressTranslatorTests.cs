using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;

namespace EPPlusTest.ExcelUtilities
{
    [TestFixture]
    public class AddressTranslatorTests
    {
        private AddressTranslator _addressTranslator;
        private ExcelDataProvider _excelDataProvider;
        private const int ExcelMaxRows = 1356;

        [SetUp]
        public void Setup()
        {
            _excelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => _excelDataProvider.ExcelMaxRows).Returns(ExcelMaxRows);
            _addressTranslator = new AddressTranslator(_excelDataProvider);
        }

        [Test]
        public void ConstructorShouldThrowIfProviderIsNull()
        {
            Assert.Throws<ArgumentNullException>(() =>
            {
                new AddressTranslator(null);
            });
        }

        [Test]
        public void ShouldTranslateRowNumber()
        {
            int col, row;
            _addressTranslator.ToColAndRow("A2", out col, out row);
            Assert.That(2, Is.EqualTo(row));
        }

        [Test]
        public void ShouldTranslateLettersToColumnIndex()
        {
            int col, row;
            _addressTranslator.ToColAndRow("C1", out col, out row);
            Assert.That(3, Is.EqualTo(col));
            _addressTranslator.ToColAndRow("AA2", out col, out row);
            Assert.That(27, Is.EqualTo(col));
            _addressTranslator.ToColAndRow("BC1", out col, out row);
            Assert.That(55, Is.EqualTo(col));
        }

        [Test]
        public void ShouldTranslateLetterAddressUsingMaxRowsFromProviderLower()
        {
            int col, row;
            _addressTranslator.ToColAndRow("A", out col, out row);
            Assert.That(1, Is.EqualTo(row));
        }

        [Test]
        public void ShouldTranslateLetterAddressUsingMaxRowsFromProviderUpper()
        {
            int col, row;
            _addressTranslator.ToColAndRow("A", out col, out row, AddressTranslator.RangeCalculationBehaviour.LastPart);
            Assert.That(ExcelMaxRows, Is.EqualTo(row));
        }
    }
}
