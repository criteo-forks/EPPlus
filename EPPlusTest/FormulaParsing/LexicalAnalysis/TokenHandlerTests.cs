using System;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestFixture]
    public class TokenHandlerTests
    {
        private TokenizerContext _tokenizerContext;
        private TokenHandler _handler;

        [SetUp]
        public void Init()
        {
            _tokenizerContext = new TokenizerContext("test");
            InitHandler(_tokenizerContext);
        }

        private void InitHandler(TokenizerContext context)
        {
            var parsingContext = ParsingContext.Create();
            var tokenFactory = new TokenFactory(parsingContext.Configuration.FunctionRepository, null);
            _handler = new TokenHandler(_tokenizerContext, tokenFactory, new TokenSeparatorProvider()); 
        }

        [Test]
        public void HasMoreTokensShouldBeTrueWhenTokensExists()
        {
            Assert.That(_handler.HasMore());
        }

        [Test]
        public void HasMoreTokensShouldBeFalseWhenAllAreHandled()
        {
            for (var x = 0; x < "test".Length; x++ )
            {
                _handler.Next();
            }
            Assert.That(!_handler.HasMore());
        }
    }
}
