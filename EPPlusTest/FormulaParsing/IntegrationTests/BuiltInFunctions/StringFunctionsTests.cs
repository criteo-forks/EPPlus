using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestFixture]
    public class StringFunctionsTests : FormulaParserTestBase
    {
        private ExcelDataProvider _provider;
        [SetUp]
        public void Setup()
        {
            _provider = A.Fake<ExcelDataProvider>();
            _parser = new FormulaParser(_provider);
        }

        [Test]
        public void TextShouldConcatenateWithNextExpression()
        {
            A.CallTo(() =>_provider.GetFormat(23.5, "$0.00")).Returns("$23.50");
            var result = _parser.Parse("TEXT(23.5,\"$0.00\") & \" per hour\"");
            Assert.That("$23.50 per hour", Is.EqualTo(result));
        }

        [Test]
        public void LenShouldAddLengthUsingSuppliedOperator()
        {
            var result = _parser.Parse("Len(\"abc\") + 2");
            Assert.That(5d, Is.EqualTo(result));
        }

        [Test]
        public void LowerShouldReturnALowerCaseString()
        {
            var result = _parser.Parse("Lower(\"ABC\")");
            Assert.That("abc", Is.EqualTo(result));
        }

        [Test]
        public void UpperShouldReturnAnUpperCaseString()
        {
            var result = _parser.Parse("Upper(\"abc\")");
            Assert.That("ABC", Is.EqualTo(result));
        }

        [Test]
        public void LeftShouldReturnSubstringFromLeft()
        {
            var result = _parser.Parse("Left(\"abacd\", 2)");
            Assert.That("ab", Is.EqualTo(result));
        }

        [Test]
        public void RightShouldReturnSubstringFromRight()
        {
            var result = _parser.Parse("RIGHT(\"abacd\", 2)");
            Assert.That("cd", Is.EqualTo(result));
        }

        [Test]
        public void MidShouldReturnSubstringAccordingToParams()
        {
            var result = _parser.Parse("Mid(\"abacd\", 2, 2)");
            Assert.That("ba", Is.EqualTo(result));
        }

        [Test]
        public void ReplaceShouldReturnSubstringAccordingToParams()
        {
            var result = _parser.Parse("Replace(\"testar\", 3, 3, \"hej\")");
            Assert.That("tehejr", Is.EqualTo(result));
        }

        [Test]
        public void SubstituteShouldReturnSubstringAccordingToParams()
        {
            var result = _parser.Parse("Substitute(\"testar testar\", \"es\", \"xx\")");
            Assert.That("txxtar txxtar", Is.EqualTo(result));
        }

        [Test]
        public void ConcatenateShouldReturnAccordingToParams()
        {
            var result = _parser.Parse("CONCATENATE(\"One\", \"Two\", \"Three\")");
            Assert.That("OneTwoThree", Is.EqualTo(result));
        }

        [Test]
        public void TShouldReturnText()
        {
            var result = _parser.Parse("T(\"One\")");
            Assert.That("One", Is.EqualTo(result));
        }

        [Test]
        public void ReptShouldConcatenate()
        {
            var result = _parser.Parse("REPT(\"*\",3)");
            Assert.That("***", Is.EqualTo(result));
        }
    }
}
