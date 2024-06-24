using System;
using System.Text;
using System.Collections.Generic;
//using System.Diagnostics.Design;
using System.Globalization;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.ExcelUtilities;

namespace EPPlusTest
{
    [TestFixture]
    public class ExpressionEvaluatorTests
    {
        private ExpressionEvaluator _evaluator;

        [SetUp]
        public void Setup()
        {
            _evaluator = new ExpressionEvaluator();
        }

        #region Numeric Expression Tests
        [Test]
        public void EvaluateShouldReturnTrueIfOperandsAreEqual()
        {
            var result = _evaluator.Evaluate("1", "1");
            Assert.That(result);
        }

        [Test]
        public void EvaluateShouldReturnTrueIfOperandsAreMatchingButDifferentTypes()
        {
            var result = _evaluator.Evaluate(1d, "1");
            Assert.That(result);
        }

        [Test]
        public void EvaluateShouldEvaluateOperator()
        {
            var result = _evaluator.Evaluate(1d, "<2");
            Assert.That(result);
        }

        [Test]
        public void EvaluateShouldEvaluateNumericString()
        {
            var result = _evaluator.Evaluate("1", ">0");
            Assert.That(result);
        }

        [Test]
        public void EvaluateShouldHandleBooleanArg()
        {
            var result = _evaluator.Evaluate(true, "TRUE");
            Assert.That(result);
        }

        [Test]
        public void EvaluateShouldThrowIfOperatorIsNotBoolean()
        {
            Assert.Throws<ArgumentException>(() =>
            {
                var result = _evaluator.Evaluate(1d, "+1");
            });
        }
        #endregion

        #region Date tests
        [Test]
        public void EvaluateShouldHandleDateArg()
        {
            #if (!Core)
                Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
            #endif
            var result = _evaluator.Evaluate(new DateTime(2016,6,28), "2016-06-28");
            Assert.That(result);
        }

        [Test]
        public void EvaluateShouldHandleDateArgWithOperator()
        {
#if (!Core)
            Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture;
#endif
            var result = _evaluator.Evaluate(new DateTime(2016, 6, 28), ">2016-06-27");
            Assert.That(result);
        }
#endregion

#region Blank Expression Tests
        [Test]
        public void EvaluateBlankExpressionEqualsNull()
        {
            var result = _evaluator.Evaluate(null, "");
            Assert.That(result);
        }

        [Test]
        public void EvaluateBlankExpressionEqualsEmptyString()
        {
            var result = _evaluator.Evaluate(string.Empty, "");
            Assert.That(!result);
        }

        [Test]
        public void EvaluateBlankExpressionEqualsZero()
        {
            var result = _evaluator.Evaluate(0d, "");
            Assert.That(!result);
        }
#endregion

#region Quotes Expression Tests
        [Test]
        public void EvaluateQuotesExpressionEqualsNull()
        {
            var result = _evaluator.Evaluate(null, "\"\"");
            Assert.That(!result);
        }

        [Test]
        public void EvaluateQuotesExpressionEqualsZero()
        {
            var result = _evaluator.Evaluate(0d, "\"\"");
            Assert.That(!result);
        }

        [Test]
        public void EvaluateQuotesExpressionEqualsCharacter()
        {
            var result = _evaluator.Evaluate("a", "\"\"");
            Assert.That(!result);
        }
#endregion

#region NotEqualToZero Expression Tests
        [Test]
        public void EvaluateNotEqualToZeroExpressionEqualsNull()
        {
            var result = _evaluator.Evaluate(null, "<>0");
            Assert.That(result);
        }

        [Test]
        public void EvaluateNotEqualToZeroExpressionEqualsEmptyString()
        {
            var result = _evaluator.Evaluate(string.Empty, "<>0");
            Assert.That(result);
        }

        [Test]
        public void EvaluateNotEqualToZeroExpressionEqualsCharacter()
        {
            var result = _evaluator.Evaluate("a", "<>0");
            Assert.That(result);
        }

        [Test]
        public void EvaluateNotEqualToZeroExpressionEqualsNonZero()
        {
            var result = _evaluator.Evaluate(1d, "<>0");
            Assert.That(result);
        }

        [Test]
        public void EvaluateNotEqualToZeroExpressionEqualsZero()
        {
            var result = _evaluator.Evaluate(0d, "<>0");
            Assert.That(!result);
        }
#endregion

#region NotEqualToBlank Expression Tests
        [Test]
        public void EvaluateNotEqualToBlankExpressionEqualsNull()
        {
            var result = _evaluator.Evaluate(null, "<>");
            Assert.That(!result);
        }

        [Test]
        public void EvaluateNotEqualToBlankExpressionEqualsEmptyString()
        {
            var result = _evaluator.Evaluate(string.Empty, "<>");
            Assert.That(result);
        }

        [Test]
        public void EvaluateNotEqualToBlankExpressionEqualsCharacter()
        {
            var result = _evaluator.Evaluate("a", "<>");
            Assert.That(result);
        }

        [Test]
        public void EvaluateNotEqualToBlankExpressionEqualsNonZero()
        {
            var result = _evaluator.Evaluate(1d, "<>");
            Assert.That(result);
        }

        [Test]
        public void EvaluateNotEqualToBlankExpressionEqualsZero()
        {
            var result = _evaluator.Evaluate(0d, "<>");
            Assert.That(result);
        }
#endregion

#region Character Expression Tests
        [Test]
        public void EvaluateCharacterExpressionEqualNull()
        {
            var result = _evaluator.Evaluate(null, "a");
            Assert.That(!result);
        }

        [Test]
        public void EvaluateCharacterExpressionEqualsEmptyString()
        {
            var result = _evaluator.Evaluate(string.Empty, "a");
            Assert.That(!result);
        }

        [Test]
        public void EvaluateCharacterExpressionEqualsNumeral()
        {
            var result = _evaluator.Evaluate(1d, "a");
            Assert.That(!result);
        }

        [Test]
        public void EvaluateCharacterExpressionEqualsSameCharacter()
        {
            var result = _evaluator.Evaluate("a", "a");
            Assert.That(result);
        }

        [Test]
        public void EvaluateCharacterExpressionEqualsDifferentCharacter()
        {
            var result = _evaluator.Evaluate("b", "a");
            Assert.That(!result);
        }
#endregion

#region CharacterWithOperator Expression Tests
        [Test]
        public void EvaluateCharacterWithOperatorExpressionEqualNull()
        {
            var result = _evaluator.Evaluate(null, ">a");
            Assert.That(!result);
            result = _evaluator.Evaluate(null, "<a");
            Assert.That(!result);
        }

        [Test]
        public void EvaluateCharacterWithOperatorExpressionEqualsEmptyString()
        {
            var result = _evaluator.Evaluate(string.Empty, ">a");
            Assert.That(!result);
            result = _evaluator.Evaluate(string.Empty, "<a");
            Assert.That(result);
        }

        [Test]
        public void EvaluateCharacterWithOperatorExpressionEqualsNumeral()
        {
            var result = _evaluator.Evaluate(1d, ">a");
            Assert.That(!result);
            result = _evaluator.Evaluate(1d, "<a");
            Assert.That(!result);
        }

        [Test]
        public void EvaluateCharacterWithOperatorExpressionEqualsSameCharacter()
        {
            var result = _evaluator.Evaluate("a", ">a");
            Assert.That(!result);
            result = _evaluator.Evaluate("a", ">=a");
            Assert.That(result);
            result = _evaluator.Evaluate("a", "<a");
            Assert.That(!result);
            result = _evaluator.Evaluate("a", ">=a");
            Assert.That(result);
        }

        [Test]
        public void EvaluateCharacterWithOperatorExpressionEqualsDifferentCharacter()
        {
            var result = _evaluator.Evaluate("b", ">a");
            Assert.That(result);
            result = _evaluator.Evaluate("b", "<a");
            Assert.That(!result);
        }
#endregion
    }
}
