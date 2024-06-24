using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing;


namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestFixture]
    public class NegationTests
    {
        private SourceCodeTokenizer _tokenizer;

        [SetUp]
        public void Setup()
        {
            var context = ParsingContext.Create();
            _tokenizer = new SourceCodeTokenizer(context.Configuration.FunctionRepository, null);
        }

        [TearDown]
        public void Cleanup()
        {

        }

        [Test]
        public void ShouldSetNegatorOnFirstTokenIfFirstCharIsMinus()
        {
            var input = "-1";
            var tokens = _tokenizer.Tokenize(input);

            Assert.That(2, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Negator, Is.EqualTo(tokens.First().TokenType));
        }

        [Test]
        public void ShouldChangePlusToMinusIfNegatorIsPresent()
        {
            var input = "1 + -1";
            var tokens = _tokenizer.Tokenize(input);

            Assert.That(3, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Operator, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That("-", Is.EqualTo(tokens.ElementAt(1).Value));
        }

        [Test]
        public void ShouldSetNegatorOnTokenInsideParenthethis()
        {
            var input = "1 + (-1 * 2)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.That(8, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Negator, Is.EqualTo(tokens.ElementAt(3).TokenType));
        }

        [Test]
        public void ShouldSetNegatorOnTokenInsideFunctionCall()
        {
            var input = "Ceiling(-1, -0.1)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.That(8, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Negator, Is.EqualTo(tokens.ElementAt(2).TokenType));
            Assert.That(TokenType.Negator, Is.EqualTo(tokens.ElementAt(5).TokenType), "Negator after comma was not identified");
        }

        [Test]
        public void ShouldSetNegatorOnTokenInEnumerable()
        {
            var input = "{-1}";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(TokenType.Negator, Is.EqualTo(tokens.ElementAt(1).TokenType));
        }

        [Test]
        public void ShouldSetNegatorOnExcelAddress()
        {
            var input = "-A1";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(TokenType.Negator, Is.EqualTo(tokens.ElementAt(0).TokenType));
            Assert.That(TokenType.ExcelAddress, Is.EqualTo(tokens.ElementAt(1).TokenType));
        }
    }
}
