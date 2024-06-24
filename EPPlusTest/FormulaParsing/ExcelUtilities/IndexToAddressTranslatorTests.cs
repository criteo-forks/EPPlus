using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using FakeItEasy;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.ExcelUtilities
{
    [TestFixture]
    public class IndexToAddressTranslatorTests
    {
        private ExcelDataProvider _excelDataProvider;
        private IndexToAddressTranslator _indexToAddressTranslator;

        [SetUp]
        public void Setup()
        {
            SetupTranslator(12345, ExcelReferenceType.RelativeRowAndColumn);
        }

        private void SetupTranslator(int maxRows, ExcelReferenceType refType)
        {
            _excelDataProvider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => _excelDataProvider.ExcelMaxRows).Returns(maxRows);
            _indexToAddressTranslator = new IndexToAddressTranslator(_excelDataProvider, refType);
        }

        [Test]
        public void ShouldThrowIfExcelDataProviderIsNull()
        {
            Assert.Throws<ArgumentNullException>(() =>
            {
                new IndexToAddressTranslator(null);
            });
        }

        [Test]
        public void ShouldTranslate1And1ToA1()
        {
            var result = _indexToAddressTranslator.ToAddress(1, 1);
            Assert.That("A1", Is.EqualTo(result));
        }

        [Test]
        public void ShouldTranslate27And1ToAA1()
        {
            var result = _indexToAddressTranslator.ToAddress(27, 1);
            Assert.That("AA1", Is.EqualTo(result));
        }

        [Test]
        public void ShouldTranslate53And1ToBA1()
        {
            var result = _indexToAddressTranslator.ToAddress(53, 1);
            Assert.That("BA1", Is.EqualTo(result));
        }

        [Test]
        public void ShouldTranslate702And1ToZZ1()
        {
            var result = _indexToAddressTranslator.ToAddress(702, 1);
            Assert.That("ZZ1", Is.EqualTo(result));
        }

        [Test]
        public void ShouldTranslate703ToAAA4()
        {
            var result = _indexToAddressTranslator.ToAddress(703, 4);
            Assert.That("AAA4", Is.EqualTo(result));
        }

        [Test]
        public void ShouldTranslateToEntireColumnWhenRowIsEqualToMaxRows()
        {
            A.CallTo(() => _excelDataProvider.ExcelMaxRows).Returns(123456);
            var result = _indexToAddressTranslator.ToAddress(1, 123456);
            Assert.That("A", Is.EqualTo(result));
        }

        [Test]
        public void ShouldTranslateToAbsoluteAddress()
        {
            SetupTranslator(123456, ExcelReferenceType.AbsoluteRowAndColumn);
            var result = _indexToAddressTranslator.ToAddress(1, 1);
            Assert.That("$A$1", Is.EqualTo(result));
        }

        [Test]
        public void ShouldTranslateToAbsoluteRowAndRelativeCol()
        {
            SetupTranslator(123456, ExcelReferenceType.AbsoluteRowRelativeColumn);
            var result = _indexToAddressTranslator.ToAddress(1, 1);
            Assert.That("A$1", Is.EqualTo(result));
        }

        [Test]
        public void ShouldTranslateToRelativeRowAndAbsoluteCol()
        {
            SetupTranslator(123456, ExcelReferenceType.RelativeRowAbsolutColumn);
            var result = _indexToAddressTranslator.ToAddress(1, 1);
            Assert.That("$A1", Is.EqualTo(result));
        }

        [Test]
        public void ShouldTranslateToRelativeRowAndCol()
        {
            SetupTranslator(123456, ExcelReferenceType.RelativeRowAndColumn);
            var result = _indexToAddressTranslator.ToAddress(1, 1);
            Assert.That("A1", Is.EqualTo(result));
        }
    }
}
