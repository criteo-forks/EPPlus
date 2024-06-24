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
    public class BasicCalcTests : FormulaParserTestBase
    {
        [SetUp]
        public void Setup()
        {
            var excelDataProvider = A.Fake<ExcelDataProvider>();
            _parser = new FormulaParser(excelDataProvider);
        }

        [Test]
        public void ShouldAddIntegersCorrectly()
        {
            var result = _parser.Parse("1 + 2");
            Assert.That(3d, Is.EqualTo(result));
        }

        [Test]
        public void ShouldSubtractIntegersCorrectly()
        {
            var result = _parser.Parse("2 - 1");
            Assert.That(1d, Is.EqualTo(result));
        }

        [Test]
        public void ShouldMultiplyIntegersCorrectly()
        {
            var result = _parser.Parse("2 * 3");
            Assert.That(6d, Is.EqualTo(result));
        }

        [Test]
        public void ShouldDivideIntegersCorrectly()
        {
            var result = _parser.Parse("8 / 4");
            Assert.That(2d, Is.EqualTo(result));
        }

        [Test]
        public void ShouldDivideDecimalWithIntegerCorrectly()
        {
            var result = _parser.Parse("2.5/2");
            Assert.That(1.25d, Is.EqualTo(result));
        }

        [Test]
        public void ShouldHandleExpCorrectly()
        {
            var result = _parser.Parse("2 ^ 4");
            Assert.That(16d, Is.EqualTo(result));
        }

        [Test]
        public void ShouldHandleExpWithDecimalCorrectly()
        {
            var result = _parser.Parse("2.5 ^ 2");
            Assert.That(6.25d, Is.EqualTo(result));
        }

        [Test]
        public void ShouldMultiplyDecimalWithDecimalCorrectly()
        {
            var result = _parser.Parse("2.5 * 1.5");
            Assert.That(3.75d, Is.EqualTo(result));
        }

        [Test]
        public void ThreeGreaterThanTwoShouldBeTrue()
        {
            var result = _parser.Parse("3 > 2");
            Assert.That((bool)result);
        }

        [Test]
        public void ThreeLessThanTwoShouldBeFalse()
        {
            var result = _parser.Parse("3 < 2");
            Assert.That(!(bool)result);
        }

        [Test]
        public void ThreeLessThanOrEqualToThreeShouldBeTrue()
        {
            var result = _parser.Parse("3 <= 3");
            Assert.That((bool)result);
        }

        [Test]
        public void ThreeLessThanOrEqualToTwoDotThreeShouldBeFalse()
        {
            var result = _parser.Parse("3 <= 2.3");
            Assert.That(!(bool)result);
        }

        [Test]
        public void ThreeGreaterThanOrEqualToThreeShouldBeTrue()
        {
            var result = _parser.Parse("3 >= 3");
            Assert.That((bool)result);
        }

        [Test]
        public void TwoDotTwoGreaterThanOrEqualToThreeShouldBeFalse()
        {
            var result = _parser.Parse("2.2 >= 3");
            Assert.That(!(bool)result);
        }

        [Test]
        public void TwelveAndTwelveShouldBeEqual()
        {
            var result = _parser.Parse("2=2");
            Assert.That((bool)result);
        }

        [Test]
        public void TenPercentShouldBe0Point1()
        {
            var result = _parser.Parse("10%");
            Assert.That(0.1, Is.EqualTo(result));
        }

        [Test]
        public void ShouldHandleMultiplePercentSigns()
        {
            var result = _parser.Parse("10%%");
            Assert.That(0.001, Is.EqualTo(result));
        }

        [Test]
        public void ShouldHandlePercentageOnFunctionResult()
        {
            var result = _parser.Parse("SUM(1;2;3)%");
            Assert.That(0.06, Is.EqualTo(result));
        }

        [Test]
        public void ShouldHandlePercentageOnParantethis()
        {
            var result = _parser.Parse("(1+2)%");
            Assert.That(0.03, Is.EqualTo(result));
        }

        [Test]
        public void ShouldIgnoreLeadingPlus()
        {
            var result = _parser.Parse("+(1-2)");
            Assert.That(-1d, Is.EqualTo(result));
        }

        [Test]
        public void ShouldHandleDecimalNumberWhenDividingIntegers()
        {
            var result = _parser.Parse("224567455/400000000*500000");
            Assert.That(280709.31875, Is.EqualTo(result));
        }

        [Test]
        public void ShouldNegateExpressionInParenthesis()
        {
            var result = _parser.Parse("-(1+2)");
            Assert.That(-3d, Is.EqualTo(result));
        }
    }
}
