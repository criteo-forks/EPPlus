using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using System.Globalization;
using System.Threading;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestFixture]
    public class ExpressionConverterTests
    {
        private IExpressionConverter _converter;

        [SetUp]
        public void Setup()
        {
            _converter = new ExpressionConverter();
        }

        [Test]
        public void ToStringExpressionShouldConvertIntegerExpressionToStringExpression()
        {
            var integerExpression = new IntegerExpression("2");
            var result = _converter.ToStringExpression(integerExpression);
            Assert.That(result, Is.InstanceOf<StringExpression>());
            Assert.That("2", Is.EqualTo(result.Compile().Result));
        }

        [Test]
        public void ToStringExpressionShouldCopyOperatorToStringExpression()
        {
            var integerExpression = new IntegerExpression("2");
            integerExpression.Operator = Operator.Plus;
            var result = _converter.ToStringExpression(integerExpression);
            Assert.That(integerExpression.Operator, Is.EqualTo(result.Operator));
        }

        [Test]
        public void ToStringExpressionShouldConvertDecimalExpressionToStringExpression()
        {
            var decimalExpression = new DecimalExpression("2.5");
            var result = _converter.ToStringExpression(decimalExpression);
            Assert.That(result, Is.InstanceOf<StringExpression>());
            Assert.That($"2{CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator}5", Is.EqualTo(result.Compile().Result));
        }

        [Test]
        public void FromCompileResultShouldCreateIntegerExpressionIfCompileResultIsInteger()
        {
            var compileResult = new CompileResult(1, DataType.Integer);
            var result = _converter.FromCompileResult(compileResult);
            Assert.That(result, Is.InstanceOf<IntegerExpression>());
            Assert.That(1d, Is.EqualTo(result.Compile().Result));
        }

        [Test]
        public void FromCompileResultShouldCreateStringExpressionIfCompileResultIsString()
        {
            var compileResult = new CompileResult("abc", DataType.String);
            var result = _converter.FromCompileResult(compileResult);
            Assert.That(result, Is.InstanceOf<StringExpression>());
            Assert.That("abc", Is.EqualTo(result.Compile().Result));
        }

        [Test]
        public void FromCompileResultShouldCreateDecimalExpressionIfCompileResultIsDecimal()
        {
            var compileResult = new CompileResult(2.5d, DataType.Decimal);
            var result = _converter.FromCompileResult(compileResult);
            Assert.That(result, Is.InstanceOf<DecimalExpression>());
            Assert.That(2.5d, Is.EqualTo(result.Compile().Result));
        }

        [Test]
        public void FromCompileResultShouldCreateBooleanExpressionIfCompileResultIsBoolean()
        {
            var compileResult = new CompileResult("true", DataType.Boolean);
            var result = _converter.FromCompileResult(compileResult);
            Assert.That(result, Is.InstanceOf<BooleanExpression>());
            Assert.That((bool)result.Compile().Result);
        }
    }
}
