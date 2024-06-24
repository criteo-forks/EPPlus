using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using FakeItEasy;

namespace EPPlusTest.ExcelUtilities
{
    [TestFixture]
    public class CellReferenceProviderTests
    {
        private ExcelDataProvider _provider;

        [SetUp]
        public void Setup()
        {
            _provider = A.Fake<ExcelDataProvider>();
            A.CallTo(() => _provider.ExcelMaxRows).Returns(5000);
        }

        [Test]
        public void ShouldReturnReferencedSingleAddress()
        {
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);
            parsingContext.Configuration.SetLexer(new Lexer(parsingContext.Configuration.FunctionRepository, parsingContext.NameValueProvider));
            parsingContext.RangeAddressFactory = new RangeAddressFactory(_provider);
            var provider = new CellReferenceProvider();
            var result = provider.GetReferencedAddresses("A1", parsingContext);
            Assert.That("A1", Is.EqualTo(result.First()));
        }

        [Test]
        public void ShouldReturnReferencedMultipleAddresses()
        {
            var parsingContext = ParsingContext.Create();
            parsingContext.Scopes.NewScope(RangeAddress.Empty);
            parsingContext.Configuration.SetLexer(new Lexer(parsingContext.Configuration.FunctionRepository, parsingContext.NameValueProvider));
            parsingContext.RangeAddressFactory = new RangeAddressFactory(_provider);
            var provider = new CellReferenceProvider();
            var result = provider.GetReferencedAddresses("A1:A2", parsingContext);
            Assert.That("A1", Is.EqualTo(result.First()));
            Assert.That("A2", Is.EqualTo(result.Last()));
        }
    }
}
