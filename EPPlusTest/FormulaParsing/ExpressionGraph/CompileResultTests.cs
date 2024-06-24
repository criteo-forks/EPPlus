using System;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
	[TestFixture]
	public class CompileResultTests
	{
		[Test]
		public void NumericStringCompileResult()
		{
			var expected = 124.24;
			string numericString = expected.ToString("n");
			CompileResult result = new CompileResult(numericString, DataType.String);
			Assert.That(!result.IsNumeric);
			Assert.That(result.IsNumericString);
			Assert.That(expected, Is.EqualTo(result.ResultNumeric));
		}

		[Test]
		public void DateStringCompileResult()
		{
			var expected = new DateTime(2013, 1, 15);
			string dateString = expected.ToString("d");
			CompileResult result = new CompileResult(dateString, DataType.String);
			Assert.That(!result.IsNumeric);
			Assert.That(result.IsDateString);
			Assert.That(expected.ToOADate(), Is.EqualTo(result.ResultNumeric));
		}
	}
}
