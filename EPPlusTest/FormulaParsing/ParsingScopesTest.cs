using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing
{
    [TestFixture]
    public class ParsingScopesTest
    {
        private ParsingScopes _parsingScopes;
        private IParsingLifetimeEventHandler _lifeTimeEventHandler;

        [SetUp]
        public void Setup()
        {
            _lifeTimeEventHandler = A.Fake<IParsingLifetimeEventHandler>();
            _parsingScopes = new ParsingScopes(_lifeTimeEventHandler);
        }

        [Test]
        public void CreatedScopeShouldBeCurrentScope()
        {
            using (var scope = _parsingScopes.NewScope(RangeAddress.Empty))
            {
                Assert.That(_parsingScopes.Current, Is.EqualTo(scope));
            }
        }

        [Test]
        public void CurrentScopeShouldHandleNestedScopes()
        {
            using (var scope1 = _parsingScopes.NewScope(RangeAddress.Empty))
            {
                Assert.That(_parsingScopes.Current, Is.EqualTo(scope1));
                using (var scope2 = _parsingScopes.NewScope(RangeAddress.Empty))
                {
                    Assert.That(_parsingScopes.Current, Is.EqualTo(scope2));
                }
                Assert.That(_parsingScopes.Current, Is.EqualTo(scope1));
            }
            Assert.That(_parsingScopes.Current, Is.Null);
        }

        [Test]
        public void CurrentScopeShouldBeNullWhenScopeHasTerminated()
        {
            using (var scope = _parsingScopes.NewScope(RangeAddress.Empty))
            { }
            Assert.That(_parsingScopes.Current, Is.Null);
        }

        [Test]
        public void NewScopeShouldSetParentOnCreatedScopeIfParentScopeExisted()
        {
            using (var scope1 = _parsingScopes.NewScope(RangeAddress.Empty))
            {
                using (var scope2 = _parsingScopes.NewScope(RangeAddress.Empty))
                {
                    Assert.That(scope1, Is.EqualTo(scope2.Parent));
                }
            }
        }

        [Test]
        public void LifetimeEventHandlerShouldBeCalled()
        {
            using (var scope = _parsingScopes.NewScope(RangeAddress.Empty))
            { }
            A.CallTo(() => _lifeTimeEventHandler.ParsingCompleted()).MustHaveHappened();
        }
    }
}