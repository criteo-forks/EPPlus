using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing
{
    [TestFixture]
    public class ParsingScopeTests
    {
        private IParsingLifetimeEventHandler _lifeTimeEventHandler;
        private ParsingScopes _parsingScopes;
        private RangeAddressFactory _factory;

        [SetUp]
        public void Setup()
        {
            var provider = A.Fake<ExcelDataProvider>();
            _factory = new RangeAddressFactory(provider);
            _lifeTimeEventHandler = A.Fake<IParsingLifetimeEventHandler>();
            _parsingScopes = A.Fake<ParsingScopes>();
        }

        [Test]
        public void ConstructorShouldSetAddress()
        {
            var expectedAddress =  _factory.Create("A1");
            var scope = new ParsingScope(_parsingScopes, expectedAddress);
            Assert.That(expectedAddress, Is.EqualTo(scope.Address));
        }

        [Test]
        public void ConstructorShouldSetParent()
        {
            var parent = new ParsingScope(_parsingScopes, _factory.Create("A1"));
            var scope = new ParsingScope(_parsingScopes, parent, _factory.Create("A2"));
            Assert.That(parent, Is.EqualTo(scope.Parent));
        }

        [Test]
        public void ScopeShouldCallKillScopeOnDispose()
        {
            var scope = new ParsingScope(_parsingScopes, _factory.Create("A1"));
            ((IDisposable)scope).Dispose();
           A.CallTo(() => _parsingScopes.KillScope(scope)).MustHaveHappened();
        }
    }
}
