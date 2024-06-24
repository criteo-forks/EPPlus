using OfficeOpenXml.FormulaParsing;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlusTest.FormulaParsing
{
    [TestFixture]
    public class ExcelAddressCacheTests
    {
        [Test]
        public void ShouldGenerateNewIds()
        {
            var cache = new ExcelAddressCache();
            var firstId = cache.GetNewId();
            Assert.That(1, Is.EqualTo(firstId));

            var secondId = cache.GetNewId();
            Assert.That(2, Is.EqualTo(secondId));
        }

        [Test]
        public void ShouldReturnCachedAddress()
        {
            var cache = new ExcelAddressCache();
            var id = cache.GetNewId();
            var address = "A1";
            var result = cache.Add(id, address);
            Assert.That(result);
            Assert.That(address, Is.EqualTo(cache.Get(id)));
        }

        [Test]
        public void AddShouldReturnFalseIfUsedId()
        {
            var cache = new ExcelAddressCache();
            var id = cache.GetNewId();
            var address = "A1";
            var result = cache.Add(id, address);
            Assert.That(result);
            var result2 = cache.Add(id, address);
            Assert.That(!result2);
        }

        [Test]
        public void ClearShouldResetId()
        {
            var cache = new ExcelAddressCache();
            var id = cache.GetNewId();
            Assert.That(1, Is.EqualTo(id));
            var address = "A1";
            var result = cache.Add(id, address);
            Assert.That(1, Is.EqualTo(cache.Count));
            var id2 = cache.GetNewId();
            Assert.That(2, Is.EqualTo(id2));
            cache.Clear();
            var id3 = cache.GetNewId();
            Assert.That(1, Is.EqualTo(id3));
            
        }
    }
}
