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
    public class ArgumentParsersTests
    {
        [Test]
        public void ShouldReturnSameInstanceOfIntParserWhenCalledTwice()
        {
            var parsers = new ArgumentParsers();
            var parser1 = parsers.GetParser(DataType.Integer);
            var parser2 = parsers.GetParser(DataType.Integer);
            Assert.That(parser1, Is.EqualTo(parser2));
        }
    }
}
