using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;
using EPPlusTest.FormulaParsing.TestHelpers;

namespace EPPlusTest.Excel.Functions
{
    [TestFixture]
    public class NumberFunctionsTests
    {
        private ParsingContext _parsingContext = ParsingContext.Create();

        [Test]
        public void CIntShouldConvertTextToInteger()
        {
            var func = new CInt();
            var args = FunctionsHelper.CreateArgs("2");
            var result = func.Execute(args, _parsingContext);
            Assert.That(2, Is.EqualTo(result.Result));
        }

        [Test]
        public void IntShouldConvertDecimalToInteger()
        {
            var func = new CInt();
            var args = FunctionsHelper.CreateArgs(2.88m);
            var result = func.Execute(args, _parsingContext);
            Assert.That(2, Is.EqualTo(result.Result));
        }

        [Test]
        public void IntShouldConvertNegativeDecimalToInteger()
        {
            var func = new CInt();
            var args = FunctionsHelper.CreateArgs(-2.88m);
            var result = func.Execute(args, _parsingContext);
            Assert.That(-3, Is.EqualTo(result.Result));
        }

        [Test]
        public void IntShouldConvertStringToInteger()
        {
            var func = new CInt();
            var args = FunctionsHelper.CreateArgs("-2.88");
            var result = func.Execute(args, _parsingContext);
            Assert.That(-3, Is.EqualTo(result.Result));
        }
    }
}
