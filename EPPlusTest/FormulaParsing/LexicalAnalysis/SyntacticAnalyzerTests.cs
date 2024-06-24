using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestFixture]
    public class SyntacticAnalyzerTests
    {
        private ISyntacticAnalyzer _analyser;

        [SetUp]
        public void Setup()
        {
            _analyser = new SyntacticAnalyzer();
        }

        [Test]
        public void ShouldPassIfParenthesisAreWellformed()
        {
            var input = new List<Token>
            {
                new Token("(", TokenType.OpeningParenthesis),
                new Token("1", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("2", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis)
            };
            _analyser.Analyze(input);
        }

        [Test]
        public void ShouldThrowExceptionIfParenthesesAreNotWellformed()
        {
            Assert.Throws<FormatException>(() =>
            {
                var input = new List<Token>
                {
                    new Token("(", TokenType.OpeningParenthesis),
                    new Token("1", TokenType.Integer),
                    new Token("+", TokenType.Operator),
                    new Token("2", TokenType.Integer)
                };
                _analyser.Analyze(input);
            });
        }

        [Test]
        public void ShouldPassIfStringIsWellformed()
        {
            var input = new List<Token>
            {
                new Token("'", TokenType.String),
                new Token("abc123", TokenType.StringContent),
                new Token("'", TokenType.String)
            };
            _analyser.Analyze(input);
        }

        [Test]
        public void ShouldThrowExceptionIfStringHasNotClosing()
        {
            Assert.Throws<FormatException>(() =>
            {
                var input = new List<Token>
                {
                    new Token("'", TokenType.String),
                    new Token("abc123", TokenType.StringContent)
                };
                _analyser.Analyze(input);
            });
        }


        [Test]
        public void ShouldThrowExceptionIfThereIsAnUnrecognizedToken()
        {
            Assert.Throws<UnrecognizedTokenException>(() =>
            {
                var input = new List<Token>
                {
                    new Token("abc123", TokenType.Unrecognized)
                };
                _analyser.Analyze(input);
            });
        }
    }
}
