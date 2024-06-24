using System;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Exceptions;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Math
{
	[TestFixture]
	public class RoundTests
	{
		[Test]
		public void RoundPositiveToOnesDownLiteral()
		{
			Round round = new Round();
			double value1 = 123.45;
		    int digits = 0;
			var result = round.Execute(new FunctionArgument[]
			{
				new FunctionArgument(value1),
				new FunctionArgument(digits)
			}, ParsingContext.Create());
			Assert.That(123D, Is.EqualTo(result.Result));
		}
        [Test]
        public void RoundPositiveToOnesUpLiteral()
        {
            Round round = new Round();
            double value1 = 123.65;
            int digits = 0;
            var result = round.Execute(new FunctionArgument[]
            {
                new FunctionArgument(value1),
                new FunctionArgument(digits)
            }, ParsingContext.Create());
            Assert.That(124D, Is.EqualTo(result.Result));
        }

        [Test]
        public void RoundPositiveToTenthsDownLiteral()
        {
            Round round = new Round();
            double value1 = 123.44;
            int digits = 1;
            var result = round.Execute(new FunctionArgument[]
            {
                new FunctionArgument(value1),
                new FunctionArgument(digits)
            }, ParsingContext.Create());
            Assert.That(123.4D, Is.EqualTo(result.Result));
        }
        [Test]
        public void RoundPositiveToTenthsUpLiteral()
        {
            Round round = new Round();
            double value1 = 123.456;
            int digits = 1;
            var result = round.Execute(new FunctionArgument[]
            {
                new FunctionArgument(value1),
                new FunctionArgument(digits)
            }, ParsingContext.Create());
            Assert.That(123.5D, Is.EqualTo(result.Result));
        }
        [Test]
        public void RoundPositiveToTensDownLiteral()
        {
            Round round = new Round();
            double value1 = 124;
            int digits = -1;
            var result = round.Execute(new FunctionArgument[]
            {
                new FunctionArgument(value1),
                new FunctionArgument(digits)
            }, ParsingContext.Create());
            Assert.That(120D, Is.EqualTo(result.Result));
        }
        [Test]
        public void RoundPositiveToTensUpLiteral()
        {
            Round round = new Round();
            double value1 = 125;
            int digits = -1;
            var result = round.Execute(new FunctionArgument[]
            {
                new FunctionArgument(value1),
                new FunctionArgument(digits)
            }, ParsingContext.Create());
            Assert.That(130D, Is.EqualTo(result.Result));
        }

        [Test]
        public void RoundNegativeToTensDownLiteral()
        {
            Round round = new Round();
            double value1 = -124;
            int digits = -1;
            var result = round.Execute(new FunctionArgument[]
            {
                new FunctionArgument(value1),
                new FunctionArgument(digits)
            }, ParsingContext.Create());
            Assert.That(-120D, Is.EqualTo(result.Result));
        }
        [Test]
        public void RoundNegativeToTensUpLiteral()
        {
            Round round = new Round();
            double value1 = -125;
            int digits = -1;
            var result = round.Execute(new FunctionArgument[]
            {
                new FunctionArgument(value1),
                new FunctionArgument(digits)
            }, ParsingContext.Create());
            Assert.That(-130D, Is.EqualTo(result.Result));
        }
        [Test]
        public void RoundNegativeToTenthsDownLiteral()
        {
            Round round = new Round();
            double value1 = -123.44;
            int digits = 1;
            var result = round.Execute(new FunctionArgument[]
            {
                new FunctionArgument(value1),
                new FunctionArgument(digits)
            }, ParsingContext.Create());
            Assert.That(-123.4D, Is.EqualTo(result.Result));
        }
        [Test]
        public void RoundNegativeToTenthsUpLiteral()
        {
            Round round = new Round();
            double value1 = -123.456;
            int digits = 1;
            var result = round.Execute(new FunctionArgument[]
            {
                new FunctionArgument(value1),
                new FunctionArgument(digits)
            }, ParsingContext.Create());
            Assert.That(-123.5D, Is.EqualTo(result.Result));
        }
        [Test]
        public void RoundNegativeMidwayLiteral()
        {
            Round round = new Round();
            double value1 = -123.5;
            int digits = 0;
            var result = round.Execute(new FunctionArgument[]
            {
                new FunctionArgument(value1),
                new FunctionArgument(digits)
            }, ParsingContext.Create());
            Assert.That(-124D, Is.EqualTo(result.Result));
        }
        [Test]
        public void RoundPositiveMidwayLiteral()
        {
            Round round = new Round();
            double value1 = 123.5;
            int digits = 0;
            var result = round.Execute(new FunctionArgument[]
            {
                new FunctionArgument(value1),
                new FunctionArgument(digits)
            }, ParsingContext.Create());
            Assert.That(124D, Is.EqualTo(result.Result));
        }
    }
}
