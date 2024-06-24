using System.Globalization;
using System.Threading;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestFixture]
    public class CompileResultFactoryTests
    {
#if (!Core)
        [Test]
        public void CalculateUsingEuropeanDates()
        {
            var us = new CultureInfo("en-US");
            Thread.CurrentThread.CurrentCulture = us;
            var crf = new CompileResultFactory();
            var result = crf.Create("1/15/2014");
            var numeric = result.ResultNumeric;
            Assert.That(41654, Is.EqualTo(numeric));
            var gb = new CultureInfo("en-GB");
            Thread.CurrentThread.CurrentCulture = gb;
            var euroResult = crf.Create("15/1/2014");
            var eNumeric = euroResult.ResultNumeric;
            Assert.That(41654, Is.EqualTo(eNumeric));
        }
#endif
    }
}
