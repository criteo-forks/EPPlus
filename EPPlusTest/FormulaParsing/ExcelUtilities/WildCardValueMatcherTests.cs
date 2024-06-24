using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.ExcelUtilities
{
    [TestFixture]
    public class WildCardValueMatcherTests
    {
        private WildCardValueMatcher _matcher;

        [SetUp]
        public void Setup()
        {
            _matcher = new WildCardValueMatcher();
        }

        [Test]
        public void IsMatchShouldReturn0WhenSingleCharWildCardMatches()
        {
            var string1 = "a?c?";
            var string2 = "abcd";
            var result = _matcher.IsMatch(string1, string2);
            Assert.That(0, Is.EqualTo(result));
        }

        [Test]
        public void IsMatchShouldReturn0WhenMultipleCharWildCardMatches()
        {
            var string1 = "a*c.";
            var string2 = "abcc.";
            var result = _matcher.IsMatch(string1, string2);
            Assert.That(0, Is.EqualTo(result));
        }
    }
}
