using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing.IntegrationTests
{
    [TestFixture]
    public class PrecedenceTests : FormulaParserTestBase
    {

        [SetUp]
        public void Setup()
        {
            var excelDataProvider = A.Fake<ExcelDataProvider>();
            _parser = new FormulaParser(excelDataProvider);
        }

        [Test]
        public void ShouldCaluclateUsingPrecedenceMultiplyBeforeAdd()
        {
            var result = _parser.Parse("4 + 6 * 2");
            Assert.That(16d, Is.EqualTo(result));
        }

        [Test]
        public void ShouldCaluclateUsingPrecedenceDivideBeforeAdd()
        {
            var result = _parser.Parse("4 + 6 / 2");
            Assert.That(7d, Is.EqualTo(result));
        }

        [Test]
        public void ShouldCalculateTwoGroupsUsingDivideAndMultiplyBeforeSubtract()
        {
            var result = _parser.Parse("4/2 + 3 * 3");
            Assert.That(11d, Is.EqualTo(result));
        }

        [Test]
        public void ShouldCalculateExpressionWithinParenthesisBeforeMultiply()
        {
            var result = _parser.Parse("(2+4) * 2");
            Assert.That(12d, Is.EqualTo(result));
        }

        [Test]
        public void ShouldConcatAfterAdd()
        {
            var result = _parser.Parse("2 + 4 & \"abc\"");
            Assert.That("6abc", Is.EqualTo(result));
        }

        [Test]
        public void Bugfixtest()
        {
            var result = _parser.Parse("(1+2)+3^2");
        }
    }
}
