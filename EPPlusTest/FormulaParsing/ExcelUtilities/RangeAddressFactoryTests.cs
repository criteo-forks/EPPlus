using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using FakeItEasy;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.ExcelUtilities
{
    [TestFixture]
    public class RangeAddressFactoryTests
    {
        private RangeAddressFactory _factory;
        private const int ExcelMaxRows = 1048576;

        [SetUp]
        public void Setup()
        {
            var provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => provider.ExcelMaxRows).Returns(ExcelMaxRows);
            _factory = new RangeAddressFactory(provider);
        }

        [Test]
        public void CreateShouldThrowIfSuppliedAddressIsNull()
        {
            Assert.Throws<ArgumentException>(() =>
            {
                _factory.Create(null);
            });
        }

        [Test]
        public void CreateShouldReturnAndInstanceWithColPropertiesSet()
        {
            var address = _factory.Create("A2");
            Assert.That(1, Is.EqualTo(address.FromCol), "FromCol was not 1");
            Assert.That(1, Is.EqualTo(address.ToCol), "ToCol was not 1");
        }

        [Test]
        public void CreateShouldReturnAndInstanceWithRowPropertiesSet()
        {
            var address = _factory.Create("A2");
            Assert.That(2, Is.EqualTo(address.FromRow), "FromRow was not 2");
            Assert.That(2, Is.EqualTo(address.ToRow), "ToRow was not 2");
        }

        [Test]
        public void CreateShouldReturnAnInstanceWithFromAndToColSetWhenARangeAddressIsSupplied()
        {
            var address = _factory.Create("A1:B2");
            Assert.That(1, Is.EqualTo(address.FromCol));
            Assert.That(2, Is.EqualTo(address.ToCol));
        }

        [Test]
        public void CreateShouldReturnAnInstanceWithFromAndToRowSetWhenARangeAddressIsSupplied()
        {
            var address = _factory.Create("A1:B3");
            Assert.That(1, Is.EqualTo(address.FromRow));
            Assert.That(3, Is.EqualTo(address.ToRow));
        }

        [Test]
        public void CreateShouldSetWorksheetNameIfSuppliedInAddress()
        {
            var address = _factory.Create("Ws!A1");
            Assert.That("Ws", Is.EqualTo(address.Worksheet));
        }

        [Test]
        public void CreateShouldReturnAnInstanceWithStringAddressSet()
        {
            var address = _factory.Create(1, 1);
            Assert.That("A1", Is.EqualTo(address.ToString()));
        }

        [Test]
        public void CreateShouldReturnAnInstanceWithFromAndToColSet()
        {
            var address = _factory.Create(1, 0);
            Assert.That(1, Is.EqualTo(address.FromCol));
            Assert.That(1, Is.EqualTo(address.ToCol));
        }

        [Test]
        public void CreateShouldReturnAnInstanceWithFromAndToRowSet()
        {
            var address = _factory.Create(0, 1);
            Assert.That(1, Is.EqualTo(address.FromRow));
            Assert.That(1, Is.EqualTo(address.ToRow));
        }

        [Test]
        public void CreateShouldReturnAnInstanceWithWorksheetSetToEmptyString()
        {
            var address = _factory.Create(0, 1);
            Assert.That(string.Empty, Is.EqualTo(address.Worksheet));
        }

        [Test]
        public void CreateShouldReturnEntireColumnRangeWhenNoRowsAreSpecified()
        {
            var address = _factory.Create("A:B");
            Assert.That(1, Is.EqualTo(address.FromCol));
            Assert.That(2, Is.EqualTo(address.ToCol));
            Assert.That(1, Is.EqualTo(address.FromRow));
            Assert.That(ExcelMaxRows, Is.EqualTo(address.ToRow));
        }
    }
}
