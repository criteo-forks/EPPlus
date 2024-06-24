using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest.ExcelUtilities
{
    [TestFixture]
    public class ValueMatcherTests
    {
        private ValueMatcher _matcher;

        [SetUp]
        public void Setup()
        {
            _matcher = new ValueMatcher();
        }

        [Test]
        public void ShouldReturn1WhenFirstParamIsSomethingAndSecondParamIsNull()
        {
            object o1 = 1;
            object o2 = null;
            var result = _matcher.IsMatch(o1, o2);
            Assert.That(1, Is.EqualTo(result));
        }

        [Test]
        public void ShouldReturnMinus1WhenFirstParamIsNullAndSecondParamIsSomething()
        {
            object o1 = null;
            object o2 = 1;
            var result = _matcher.IsMatch(o1, o2);
            Assert.That(-1, Is.EqualTo(result));
        }

        [Test]
        public void ShouldReturn0WhenBothParamsAreNull()
        {
            object o1 = null;
            object o2 = null;
            var result = _matcher.IsMatch(o1, o2);
            Assert.That(0, Is.EqualTo(result));
        }

        [Test]
        public void ShouldReturn0WhenBothParamsAreEqual()
        {
            object o1 = 1d;
            object o2 = 1d;
            var result = _matcher.IsMatch(o1, o2);
            Assert.That(0, Is.EqualTo(result));
        }

        [Test]
        public void ShouldReturnMinus1WhenFirstParamIsLessThanSecondParam()
        {
            object o1 = 1d;
            object o2 = 5d;
            var result = _matcher.IsMatch(o1, o2);
            Assert.That(-1, Is.EqualTo(result));
        }

        [Test]
        public void ShouldReturn1WhenFirstParamIsGreaterThanSecondParam()
        {
            object o1 = 3d;
            object o2 = 1d;
            var result = _matcher.IsMatch(o1, o2);
            Assert.That(1, Is.EqualTo(result));
        }

        [Test]
        public void ShouldReturn0WhenWhenParamsAreEqualStrings()
        {
            object o1 = "T";
            object o2 = "T";
            var result = _matcher.IsMatch(o1, o2);
            Assert.That(0, Is.EqualTo(result));
        }

        [Test]
        public void ShouldReturn0WhenParamsAreEqualButDifferentTypes()
        {
            object o1 = "2";
            object o2 = 2d;
            var result = _matcher.IsMatch(o1, o2);
            Assert.That(0, Is.EqualTo(result), "IsMatch did not return 0 as expected when first param is a string and second a double");

            o1 = 2d;
            o2 = "2";
            result = _matcher.IsMatch(o1, o2);
            Assert.That(0, Is.EqualTo(result), "IsMatch did not return 0 as expected when first param is a double and second a string");
        }

        [Test]
        public void ShouldReturnMînus2WhenTypesDifferAndStringConversionToDoubleFails()
        {
            object o1 = 2d;
            object o2 = "T";
            var result = _matcher.IsMatch(o1, o2);
            Assert.That(-2, Is.EqualTo(result));
        }
    }
}
