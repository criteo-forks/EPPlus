using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.Excel.Functions
{
    [TestFixture]
    public class ArgumentParserFactoryTests
    {
        private ArgumentParserFactory _parserFactory;

        [SetUp]
        public void Setup()
        {
            _parserFactory = new ArgumentParserFactory();
        }

        [Test]
        public void ShouldReturnIntArgumentParserWhenDataTypeIsInteger()
        {
            var parser = _parserFactory.CreateArgumentParser(DataType.Integer);
            Assert.That(parser, Is.InstanceOf<IntArgumentParser>());
        }

        [Test]
        public void ShouldReturnBoolArgumentParserWhenDataTypeIsBoolean()
        {
            var parser = _parserFactory.CreateArgumentParser(DataType.Boolean);
            Assert.That(parser, Is.InstanceOf<BoolArgumentParser>());
        }

        [Test]
        public void ShouldReturnDoubleArgumentParserWhenDataTypeIsDecial()
        {
            var parser = _parserFactory.CreateArgumentParser(DataType.Decimal);
            Assert.That(parser, Is.InstanceOf<DoubleArgumentParser>());
        }
    }
}
