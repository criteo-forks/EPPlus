using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.Excel.Functions.Text
{
    [TestFixture]
    public class TextFunctionsTests
    {
        private ParsingContext _parsingContext = ParsingContext.Create();

        [Test]
        public void CStrShouldConvertNumberToString()
        {
            var func = new CStr();
            var result = func.Execute(FunctionsHelper.CreateArgs(1), _parsingContext);
            Assert.That(DataType.String, Is.EqualTo(result.DataType));
            Assert.That("1", Is.EqualTo(result.Result));
        }

        [Test]
        public void LenShouldReturnStringsLength()
        {
            var func = new Len();
            var result = func.Execute(FunctionsHelper.CreateArgs("abc"), _parsingContext);
            Assert.That(3d, Is.EqualTo(result.Result));
        }

        [Test]
        public void LowerShouldReturnLowerCaseString()
        {
            var func = new Lower();
            var result = func.Execute(FunctionsHelper.CreateArgs("ABC"), _parsingContext);
            Assert.That("abc", Is.EqualTo(result.Result));
        }

        [Test]
        public void UpperShouldReturnUpperCaseString()
        {
            var func = new Upper();
            var result = func.Execute(FunctionsHelper.CreateArgs("abc"), _parsingContext);
            Assert.That("ABC", Is.EqualTo(result.Result));
        }

        [Test]
        public void LeftShouldReturnSubstringFromLeft()
        {
            var func = new Left();
            var result = func.Execute(FunctionsHelper.CreateArgs("abcd", 2), _parsingContext);
            Assert.That("ab", Is.EqualTo(result.Result));
        }

        [Test]
        public void RightShouldReturnSubstringFromRight()
        {
            var func = new Right();
            var result = func.Execute(FunctionsHelper.CreateArgs("abcd", 2), _parsingContext);
            Assert.That("cd", Is.EqualTo(result.Result));
        }

        [Test]
        public void MidShouldReturnSubstringAccordingToParams()
        {
            var func = new Mid();
            var result = func.Execute(FunctionsHelper.CreateArgs("abcd", 1, 2), _parsingContext);
            Assert.That("ab", Is.EqualTo(result.Result));
        }

        [Test]
        public void ReplaceShouldReturnAReplacedStringAccordingToParamsWhenStartIxIs1()
        {
            var func = new Replace();
            var result = func.Execute(FunctionsHelper.CreateArgs("testar", 1, 2, "hej"), _parsingContext);
            Assert.That("hejstar", Is.EqualTo(result.Result));
        }

        [Test]
        public void ReplaceShouldReturnAReplacedStringAccordingToParamsWhenStartIxIs3()
        {
            var func = new Replace();
            var result = func.Execute(FunctionsHelper.CreateArgs("testar", 3, 3, "hej"), _parsingContext);
            Assert.That("tehejr", Is.EqualTo(result.Result));
        }

        [Test]
        public void SubstituteShouldReturnAReplacedStringAccordingToParamsWhen()
        {
            var func = new Substitute();
            var result = func.Execute(FunctionsHelper.CreateArgs("testar testar", "es", "xx"), _parsingContext);
            Assert.That("txxtar txxtar", Is.EqualTo(result.Result));
        }

        [Test]
        public void ConcatenateShouldConcatenateThreeStrings()
        {
            var func = new Concatenate();
            var result = func.Execute(FunctionsHelper.CreateArgs("One", "Two", "Three"), _parsingContext);
            Assert.That("OneTwoThree", Is.EqualTo(result.Result));
        }

        [Test]
        public void ConcatenateShouldConcatenateStringWithInt()
        {
            var func = new Concatenate();
            var result = func.Execute(FunctionsHelper.CreateArgs(1, "Two"), _parsingContext);
            Assert.That("1Two", Is.EqualTo(result.Result));
        }

        [Test]
        public void ExactShouldReturnTrueWhenTwoEqualStrings()
        {
            var func = new Exact();
            var result = func.Execute(FunctionsHelper.CreateArgs("abc", "abc"), _parsingContext);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void ExactShouldReturnTrueWhenEqualStringAndDouble()
        {
            var func = new Exact();
            var result = func.Execute(FunctionsHelper.CreateArgs("1", 1d), _parsingContext);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void ExactShouldReturnFalseWhenStringAndNull()
        {
            var func = new Exact();
            var result = func.Execute(FunctionsHelper.CreateArgs("1", null), _parsingContext);
            Assert.That(!(bool)result.Result);
        }

        [Test]
        public void ExactShouldReturnFalseWhenTwoEqualStringsWithDifferentCase()
        {
            var func = new Exact();
            var result = func.Execute(FunctionsHelper.CreateArgs("abc", "Abc"), _parsingContext);
            Assert.That(!(bool)result.Result);
        }

        [Test]
        public void FindShouldReturnIndexOfFoundPhrase()
        {
            var func = new Find();
            var result = func.Execute(FunctionsHelper.CreateArgs("hopp", "hej hopp"), _parsingContext);
            Assert.That(5, Is.EqualTo(result.Result));
        }

        [Test]
        public void FindShouldReturnIndexOfFoundPhraseBasedOnStartIndex()
        {
            var func = new Find();
            var result = func.Execute(FunctionsHelper.CreateArgs("hopp", "hopp hopp", 2), _parsingContext);
            Assert.That(6, Is.EqualTo(result.Result));
        }

        [Test]
        public void ProperShouldSetFirstLetterToUpperCase()
        {
            var func = new Proper();
            var result = func.Execute(FunctionsHelper.CreateArgs("this IS A tEst.wi3th SOME w0rds östEr"), _parsingContext);
            Assert.That("This Is A Test.Wi3Th Some W0Rds Öster", Is.EqualTo(result.Result));
        }

        [Test]
        public void HyperLinkShouldReturnArgIfOneArgIsSupplied()
        {
            var func = new Hyperlink();
            var result = func.Execute(FunctionsHelper.CreateArgs("http://epplus.codeplex.com"), _parsingContext);
            Assert.That("http://epplus.codeplex.com", Is.EqualTo(result.Result));
        }

        [Test]
        public void HyperLinkShouldReturnLastArgIfTwoArgsAreSupplied()
        {
            var func = new Hyperlink();
            var result = func.Execute(FunctionsHelper.CreateArgs("http://epplus.codeplex.com", "EPPlus"), _parsingContext);
            Assert.That("EPPlus", Is.EqualTo(result.Result));
        }
    }
}
