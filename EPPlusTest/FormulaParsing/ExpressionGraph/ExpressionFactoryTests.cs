using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using FakeItEasy;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestFixture]
    public class ExpressionFactoryTests
    {
        private IExpressionFactory _factory;
        private ParsingContext _parsingContext;

        [SetUp]
        public void Setup()
        {
            _parsingContext = ParsingContext.Create();
            var provider = A.Fake<ExcelDataProvider>();
            _factory = new ExpressionFactory(provider, _parsingContext);
        }

        [Test]
        public void ShouldReturnIntegerExpressionWhenTokenIsInteger()
        {
            var token = new Token("2", TokenType.Integer);
            var expression = _factory.Create(token);
            Assert.That(expression, Is.InstanceOf<IntegerExpression>());
        }

        [Test]
        public void ShouldReturnBooleanExpressionWhenTokenIsBoolean()
        {
            var token = new Token("true", TokenType.Boolean);
            var expression = _factory.Create(token);
            Assert.That(expression, Is.InstanceOf<BooleanExpression>());
        }

        [Test]
        public void ShouldReturnDecimalExpressionWhenTokenIsDecimal()
        {
            var token = new Token("2.5", TokenType.Decimal);
            var expression = _factory.Create(token);
            Assert.That(expression, Is.InstanceOf<DecimalExpression>());
        }

        [Test]
        public void ShouldReturnExcelRangeExpressionWhenTokenIsExcelAddress()
        {
            var token = new Token("A1", TokenType.ExcelAddress);
            var expression = _factory.Create(token);
            Assert.That(expression, Is.InstanceOf<ExcelAddressExpression>());
        }

        [Test]
        public void ShouldReturnNamedValueExpressionWhenTokenIsNamedValue()
        {
            var token = new Token("NamedValue", TokenType.NameValue);
            var expression = _factory.Create(token);
            Assert.That(expression, Is.InstanceOf<NamedValueExpression>());
        }
    }
}
