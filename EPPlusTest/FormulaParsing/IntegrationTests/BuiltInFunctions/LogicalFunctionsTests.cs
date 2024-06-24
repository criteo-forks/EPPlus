using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestFixture]
    public class LogicalFunctionsTests : FormulaParserTestBase
    {
        [SetUp]
        public void Setup()
        {
            var excelDataProvider = A.Fake<ExcelDataProvider>();
            _parser = new FormulaParser(excelDataProvider);
        }

        [Test]
        public void IfShouldReturnCorrectResult()
        {
            var result = _parser.Parse("If(2 < 3, 1, 2)");
            Assert.That(1d, Is.EqualTo(result));
        }

        [Test]
        public void IIfShouldReturnCorrectResultWhenInnerFunctionExists()
        {
            var result = _parser.Parse("If(NOT(Or(true, FALSE)), 1, 2)");
            Assert.That(2d, Is.EqualTo(result));
        }

        [Test]
        public void IIfShouldReturnCorrectResultWhenTrueConditionIsCoercedFromAString()
        {
            var result = _parser.Parse(@"If(""true"", 1, 2)");
            Assert.That(1d, Is.EqualTo(result));
        }

        [Test]
        public void IIfShouldReturnCorrectResultWhenFalseConditionIsCoercedFromAString()
        {
            var result = _parser.Parse(@"If(""false"", 1, 2)");
            Assert.That(2d, Is.EqualTo(result));
        }

        [Test]
        public void NotShouldReturnCorrectResult()
        {
            var result = _parser.Parse("not(true)");
            Assert.That(!(bool)result);

            result = _parser.Parse("NOT(false)");
            Assert.That((bool)result);
        }

        [Test]
        public void AndShouldReturnCorrectResult()
        {
            var result = _parser.Parse("And(true, 1)");
            Assert.That((bool)result);

            result = _parser.Parse("AND(true, true, 1, false)");
            Assert.That(!(bool)result);
        }

        [Test]
        public void OrShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Or(FALSE, 0)");
            Assert.That(!(bool)result);

            result = _parser.Parse("OR(true, true, 1, false)");
            Assert.That((bool)result);
        }

        [Test]
        public void TrueShouldReturnCorrectResult()
        {
            var result = _parser.Parse("True()");
            Assert.That((bool)result);
        }

        [Test]
        public void FalseShouldReturnCorrectResult()
        {
            var result = _parser.Parse("False()");
            Assert.That(!(bool)result);
        }
    }
}
