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
    public class RangeAddressTests
    {
        private RangeAddressFactory _factory;

        [SetUp]
        public void Setup()
        {
            var provider = A.Fake<ExcelDataProvider>();
            _factory = new RangeAddressFactory(provider);
        }

        [Test]
        public void CollideShouldReturnTrueIfRangesCollides()
        {
            var address1 = _factory.Create("A1:A6");
            var address2 = _factory.Create("A5");
            Assert.That(address1.CollidesWith(address2));
        }

        [Test]
        public void CollideShouldReturnFalseIfRangesDoesNotCollide()
        {
            var address1 = _factory.Create("A1:A6");
            var address2 = _factory.Create("A8");
            Assert.That(!address1.CollidesWith(address2));
        }

        [Test]
        public void CollideShouldReturnFalseIfRangesCollidesButWorksheetNameDiffers()
        {
            var address1 = _factory.Create("Ws!A1:A6");
            var address2 = _factory.Create("A5");
            Assert.That(!address1.CollidesWith(address2));
        }
    }
}
