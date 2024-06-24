using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestFixture]
    public class TokenFactoryTests
    {
        private ITokenFactory _tokenFactory;
        private INameValueProvider _nameValueProvider;


        [SetUp]
        public void Setup()
        {
            var context = ParsingContext.Create();
            var excelDataProvider = A.Fake<ExcelDataProvider>();
            _nameValueProvider = A.Fake<INameValueProvider>();
            _tokenFactory = new TokenFactory(context.Configuration.FunctionRepository, _nameValueProvider);
        }

        [TearDown]
        public void Cleanup()
        {
      
        }

        [Test]
        public void ShouldCreateAStringToken()
        {
            var input = "\"";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.That("\"", Is.EqualTo(token.Value));
            Assert.That(TokenType.String, Is.EqualTo(token.TokenType));
        }

        [Test]
        public void ShouldCreatePlusAsOperatorToken()
        {
            var input = "+";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.That("+", Is.EqualTo(token.Value));
            Assert.That(TokenType.Operator, Is.EqualTo(token.TokenType));
        }

        [Test]
        public void ShouldCreateMinusAsOperatorToken()
        {
            var input = "-";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.That("-", Is.EqualTo(token.Value));
            Assert.That(TokenType.Operator, Is.EqualTo(token.TokenType));
        }

        [Test]
        public void ShouldCreateMultiplyAsOperatorToken()
        {
            var input = "*";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.That("*", Is.EqualTo(token.Value));
            Assert.That(TokenType.Operator, Is.EqualTo(token.TokenType));
        }

        [Test]
        public void ShouldCreateDivideAsOperatorToken()
        {
            var input = "/";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.That("/", Is.EqualTo(token.Value));
            Assert.That(TokenType.Operator, Is.EqualTo(token.TokenType));
        }

        [Test]
        public void ShouldCreateEqualsAsOperatorToken()
        {
            var input = "=";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.That("=", Is.EqualTo(token.Value));
            Assert.That(TokenType.Operator, Is.EqualTo(token.TokenType));
        }

        [Test]
        public void ShouldCreateIntegerAsIntegerToken()
        {
            var input = "23";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.That("23", Is.EqualTo(token.Value));
            Assert.That(TokenType.Integer, Is.EqualTo(token.TokenType));
        }

        [Test]
        public void ShouldCreateBooleanAsBooleanToken()
        {
            var input = "true";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.That("true", Is.EqualTo(token.Value));
            Assert.That(TokenType.Boolean, Is.EqualTo(token.TokenType));
        }

        [Test]
        public void ShouldCreateDecimalAsDecimalToken()
        {
            var input = "23.3";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);

            Assert.That("23.3", Is.EqualTo(token.Value));
            Assert.That(TokenType.Decimal, Is.EqualTo(token.TokenType));
        }

        [Test]
        public void CreateShouldReadFunctionsFromFuncRepository()
        {
            var input = "Text";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
            Assert.That(TokenType.Function, Is.EqualTo(token.TokenType));
            Assert.That("Text", Is.EqualTo(token.Value));
        }

        [Test]
        public void CreateShouldCreateExcelAddressAsExcelAddressToken()
        {
            var input = "A1";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
            Assert.That(TokenType.ExcelAddress, Is.EqualTo(token.TokenType));
            Assert.That("A1", Is.EqualTo(token.Value));
        }

        [Test]
        public void CreateShouldCreateExcelRangeAsExcelAddressToken()
        {
            var input = "A1:B15";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
            Assert.That(TokenType.ExcelAddress, Is.EqualTo(token.TokenType));
            Assert.That("A1:B15", Is.EqualTo(token.Value));
        }

        [Test]
        public void CreateShouldCreateExcelRangeOnOtherSheetAsExcelAddressToken()
        {
            var input = "ws!A1:B15";
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
            Assert.That(TokenType.ExcelAddress, Is.EqualTo(token.TokenType));
            Assert.That("WS!A1:B15", Is.EqualTo(token.Value));
        }

        [Test]
        public void CreateShouldCreateNamedValueAsExcelAddressToken()
        {
            var input = "NamedValue";
            A.CallTo(() => _nameValueProvider.IsNamedValue("NamedValue","")).Returns(true);
            A.CallTo(() => _nameValueProvider.IsNamedValue("NamedValue", null)).Returns(true);
            var token = _tokenFactory.Create(Enumerable.Empty<Token>(), input);
            Assert.That(TokenType.NameValue, Is.EqualTo(token.TokenType));
            Assert.That("NamedValue", Is.EqualTo(token.Value));
        }
    }
}
