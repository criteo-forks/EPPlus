using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.Excel.Functions;

namespace EPPlusTest.Excel.Functions
{
    [TestFixture]
    public class ArgumentParsersImplementationsTests
    {
        [Test]
        public void IntParserShouldThrowIfArgumentIsNull()
        {
            Assert.Throws<ArgumentNullException>(() =>
            {
                var parser = new IntArgumentParser();
                parser.Parse(null);
            });
        }

        [Test]
        public void IntParserShouldConvertToAnInteger()
        {
            var parser = new IntArgumentParser();
            var result = parser.Parse(3);
            Assert.That(3, Is.EqualTo(result));
        }

        [Test]
        public void IntParserShouldConvertADoubleToAnInteger()
        {
            var parser = new IntArgumentParser();
            var result = parser.Parse(3d);
            Assert.That(3, Is.EqualTo(result));
        }

        [Test]
        public void IntParserShouldConvertAStringValueToAnInteger()
        {
            var parser = new IntArgumentParser();
            var result = parser.Parse("3");
            Assert.That(3, Is.EqualTo(result));
        }

        [Test]
        public void BoolParserShouldConvertNullToFalse()
        {
            var parser = new BoolArgumentParser();
            var result = (bool)parser.Parse(null);
            Assert.That(!result);
        }

        [Test]
        public void BoolParserShouldConvertStringValueTrueToBoolValueTrue()
        {
            var parser = new BoolArgumentParser();
            var result = (bool)parser.Parse("true");
            Assert.That(result);
        }

        [Test]
        public void BoolParserShouldConvert0ToFalse()
        {
            var parser = new BoolArgumentParser();
            var result = (bool)parser.Parse(0);
            Assert.That(!result);
        }

        [Test]
        public void BoolParserShouldConvert1ToTrue()
        {
            var parser = new BoolArgumentParser();
            var result = (bool)parser.Parse(0);
            Assert.That(!result);
        }

        [Test]
        public void DoubleParserShouldConvertDoubleToDouble()
        {
            var parser = new DoubleArgumentParser();
            var result = parser.Parse(3d);
            Assert.That(3d, Is.EqualTo(result));
        }

        [Test]
        public void DoubleParserShouldConvertIntToDouble()
        {
            var parser = new DoubleArgumentParser();
            var result = parser.Parse(3);
            Assert.That(3d, Is.EqualTo(result));
        }

        [Test]
        public void DoubleParserShouldThrowIfArgumentIsNull()
        {
            Assert.Throws<ArgumentNullException>(() =>
            {
                var parser = new DoubleArgumentParser();
                parser.Parse(null);
            });
        }

        [Test]
        public void DoubleParserConvertStringToDoubleWithDotSeparator()
        {
            var parser = new DoubleArgumentParser();
            var result = parser.Parse("3.3");
            Assert.That(3.3d, Is.EqualTo(result));
        }

        [Test]
        public void DoubleParserConvertDateStringToDouble()
        {
            var parser = new DoubleArgumentParser();
            var result = parser.Parse("3.3.2015");
            Assert.That(new DateTime(2015,3,3).ToOADate(), Is.EqualTo(result));
        }
    }
}
