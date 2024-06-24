using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing
{
    [TestFixture]
    public class ParsingContextTests
    {
        [Test]
        public void ConfigurationShouldBeSetByFactoryMethod()
        {
            var context = ParsingContext.Create();
            Assert.That(context.Configuration, Is.Not.Null);
        }

        [Test]
        public void ScopesShouldBeSetByFactoryMethod()
        {
            var context = ParsingContext.Create();
            Assert.That(context.Scopes, Is.Not.Null);
        }
    }
}
