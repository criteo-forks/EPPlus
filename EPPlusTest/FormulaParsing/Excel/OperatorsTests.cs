using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing.Exceptions;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.Excel
{
    [TestFixture]
    public class OperatorsTests
    {
        [Test]
        public void OperatorPlusShouldThrowExceptionIfNonNumericOperand()
        {
            var result = Operator.Plus.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
            Assert.That(ExcelErrorValue.Create(eErrorType.Value), Is.EqualTo(result.Result));
        }

        [Test]
        public void OperatorPlusShouldAddNumericStringAndNumber()
        {
            var result = Operator.Plus.Apply(new CompileResult(1, DataType.Integer), new CompileResult("2", DataType.String));
            Assert.That(3d, Is.EqualTo(result.Result));
        }

        [Test]
        public void OperatorMinusShouldThrowExceptionIfNonNumericOperand()
        {
            var result = Operator.Minus.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
            Assert.That(ExcelErrorValue.Create(eErrorType.Value), Is.EqualTo(result.Result));
        }

        [Test]
        public void OperatorMinusShouldSubtractNumericStringAndNumber()
        {
            var result = Operator.Minus.Apply(new CompileResult(5, DataType.Integer), new CompileResult("2", DataType.String));
            Assert.That(3d, Is.EqualTo(result.Result));
        }

        [Test]
        public void OperatorDivideShouldReturnDivideByZeroIfRightOperandIsZero()
        {
            var result = Operator.Divide.Apply(new CompileResult(1d, DataType.Decimal), new CompileResult(0d, DataType.Decimal));
            Assert.That(ExcelErrorValue.Create(eErrorType.Div0), Is.EqualTo(result.Result));
        }

        [Test]
        public void OperatorDivideShouldDivideCorrectly()
        {
            var result = Operator.Divide.Apply(new CompileResult(9d, DataType.Decimal), new CompileResult(3d, DataType.Decimal));
            Assert.That(3d, Is.EqualTo(result.Result));
        }

        [Test]
        public void OperatorDivideShouldReturnValueErrorIfNonNumericOperand()
        {
            var result = Operator.Divide.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
            Assert.That(ExcelErrorValue.Create(eErrorType.Value), Is.EqualTo(result.Result));
        }

        [Test]
        public void OperatorDivideShouldDivideNumericStringAndNumber()
        {
            var result = Operator.Divide.Apply(new CompileResult(9, DataType.Integer), new CompileResult("3", DataType.String));
            Assert.That(3d, Is.EqualTo(result.Result));
        }

        [Test]
        public void OperatorMultiplyShouldThrowExceptionIfNonNumericOperand()
        {
            Operator.Multiply.Apply(new CompileResult(1, DataType.Integer), new CompileResult("a", DataType.String));
        }

        [Test]
        public void OperatoMultiplyShouldMultiplyNumericStringAndNumber()
        {
            var result = Operator.Multiply.Apply(new CompileResult(1, DataType.Integer), new CompileResult("3", DataType.String));
            Assert.That(3d, Is.EqualTo(result.Result));
        }

        [Test]
        public void OperatorConcatShouldConcatTwoStrings()
        {
            var result = Operator.Concat.Apply(new CompileResult("a", DataType.String), new CompileResult("b", DataType.String));
            Assert.That("ab", Is.EqualTo(result.Result));
        }

        [Test]
        public void OperatorConcatShouldConcatANumberAndAString()
        {
            var result = Operator.Concat.Apply(new CompileResult(12, DataType.Integer), new CompileResult("b", DataType.String));
            Assert.That("12b", Is.EqualTo(result.Result));
        }

        [Test]
        public void OperatorEqShouldReturnTruefSuppliedValuesAreEqual()
        {
            var result = Operator.Eq.Apply(new CompileResult(12, DataType.Integer), new CompileResult(12, DataType.Integer));
            Assert.That((bool)result.Result);
        }

        [Test]
        public void OperatorEqShouldReturnFalsefSuppliedValuesDiffer()
        {
            var result = Operator.Eq.Apply(new CompileResult(11, DataType.Integer), new CompileResult(12, DataType.Integer));
            Assert.That(!(bool)result.Result);
        }

        [Test]
        public void OperatorNotEqualToShouldReturnTruefSuppliedValuesDiffer()
        {
            var result = Operator.NotEqualsTo.Apply(new CompileResult(11, DataType.Integer), new CompileResult(12, DataType.Integer));
            Assert.That((bool)result.Result);
        }

        [Test]
        public void OperatorNotEqualToShouldReturnFalsefSuppliedValuesAreEqual()
        {
            var result = Operator.NotEqualsTo.Apply(new CompileResult(11, DataType.Integer), new CompileResult(11, DataType.Integer));
            Assert.That(!(bool)result.Result);
        }

        [Test]
        public void OperatorGreaterThanToShouldReturnTrueIfLeftIsSetAndRightIsNull()
        {
            var result = Operator.GreaterThan.Apply(new CompileResult(11, DataType.Integer), new CompileResult(null, DataType.Empty));
            Assert.That((bool)result.Result);
        }

        [Test]
        public void OperatorGreaterThanToShouldReturnTrueIfLeftIs11AndRightIs10()
        {
            var result = Operator.GreaterThan.Apply(new CompileResult(11, DataType.Integer), new CompileResult(10, DataType.Integer));
            Assert.That((bool)result.Result);
        }

        [Test]
        public void OperatorExpShouldReturnCorrectResult()
        {
            var result = Operator.Exp.Apply(new CompileResult(2, DataType.Integer), new CompileResult(3, DataType.Integer));
            Assert.That(8d, Is.EqualTo(result.Result));
        }

        [Test]
		public void OperatorsActingOnNumericStrings()
		{
			double number1 = 42.0;
			double number2 = -143.75;
			CompileResult result1 = new CompileResult(number1.ToString("n"), DataType.String);
			CompileResult result2 = new CompileResult(number2.ToString("n"), DataType.String);
			var operatorResult = Operator.Concat.Apply(result1, result2);
			Assert.That($"{number1.ToString("n")}{number2.ToString("n")}", Is.EqualTo(operatorResult.Result));
			operatorResult = Operator.Divide.Apply(result1, result2);
			Assert.That(number1 / number2, Is.EqualTo(operatorResult.Result));
			operatorResult = Operator.Exp.Apply(result1, result2);
			Assert.That(Math.Pow(number1, number2), Is.EqualTo(operatorResult.Result));
			operatorResult = Operator.Minus.Apply(result1, result2);
			Assert.That(number1 - number2, Is.EqualTo(operatorResult.Result));
			operatorResult = Operator.Multiply.Apply(result1, result2);
			Assert.That(number1 * number2, Is.EqualTo(operatorResult.Result));
			operatorResult = Operator.Percent.Apply(result1, result2);
			Assert.That(number1 * number2, Is.EqualTo(operatorResult.Result));
			operatorResult = Operator.Plus.Apply(result1, result2);
			Assert.That(number1 + number2, Is.EqualTo(operatorResult.Result));
			// Comparison operators always compare string-wise and don't parse out the actual numbers.
			operatorResult = Operator.NotEqualsTo.Apply(result1, new CompileResult(number1.ToString("n0"), DataType.String));
			Assert.That((bool)operatorResult.Result);
			operatorResult = Operator.Eq.Apply(result1, new CompileResult(number1.ToString("n0"), DataType.String));
			Assert.That(!(bool)operatorResult.Result);
			operatorResult = Operator.GreaterThan.Apply(result1, result2);
			Assert.That((bool)operatorResult.Result);
			operatorResult = Operator.GreaterThanOrEqual.Apply(result1, result2);
			Assert.That((bool)operatorResult.Result);
			operatorResult = Operator.LessThan.Apply(result1, result2);
			Assert.That(!(bool)operatorResult.Result);
			operatorResult = Operator.LessThanOrEqual.Apply(result1, result2);
			Assert.That(!(bool)operatorResult.Result);
		}

		[Test]
		public void OperatorsActingOnDateStrings()
		{
            const string dateFormat = "M-dd-yyyy";
            DateTime date1 = new DateTime(2015, 2, 20);
            DateTime date2 = new DateTime(2015, 12, 1);
            var numericDate1 = date1.ToOADate();
            var numericDate2 = date2.ToOADate();
            CompileResult result1 = new CompileResult(date1.ToString(dateFormat), DataType.String); // 2/20/2015
            CompileResult result2 = new CompileResult(date2.ToString(dateFormat), DataType.String); // 12/1/2015
            var operatorResult = Operator.Concat.Apply(result1, result2);
            Assert.That($"{date1.ToString(dateFormat)}{date2.ToString(dateFormat)}", Is.EqualTo(operatorResult.Result));
            operatorResult = Operator.Divide.Apply(result1, result2);
            Assert.That(numericDate1 / numericDate2, Is.EqualTo(operatorResult.Result));
            operatorResult = Operator.Exp.Apply(result1, result2);
            Assert.That(Math.Pow(numericDate1, numericDate2), Is.EqualTo(operatorResult.Result));
            operatorResult = Operator.Minus.Apply(result1, result2);
			Assert.That(numericDate1 - numericDate2, Is.EqualTo(operatorResult.Result));
			operatorResult = Operator.Multiply.Apply(result1, result2);
			Assert.That(numericDate1 * numericDate2, Is.EqualTo(operatorResult.Result));
			operatorResult = Operator.Percent.Apply(result1, result2);
			Assert.That(numericDate1 * numericDate2, Is.EqualTo(operatorResult.Result));
			operatorResult = Operator.Plus.Apply(result1, result2);
			Assert.That(numericDate1 + numericDate2, Is.EqualTo(operatorResult.Result));
			// Comparison operators always compare string-wise and don't parse out the actual numbers.
			operatorResult = Operator.Eq.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String));
			Assert.That(!(bool)operatorResult.Result);
			operatorResult = Operator.NotEqualsTo.Apply(result1, new CompileResult(date1.ToString("f"), DataType.String));
			Assert.That((bool)operatorResult.Result);
			operatorResult = Operator.GreaterThan.Apply(result1, result2);
			Assert.That((bool)operatorResult.Result);
			operatorResult = Operator.GreaterThanOrEqual.Apply(result1, result2);
			Assert.That((bool)operatorResult.Result);
			operatorResult = Operator.LessThan.Apply(result1, result2);
			Assert.That(!(bool)operatorResult.Result);
			operatorResult = Operator.LessThanOrEqual.Apply(result1, result2);
			Assert.That(!(bool)operatorResult.Result);
		}
	}
}
