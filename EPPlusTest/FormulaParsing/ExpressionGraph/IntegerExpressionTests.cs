using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Operators;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestFixture]
    public class IntegerExpressionTests
    {
        [Test]
        public void MergeWithNextWithPlusOperatorShouldCalulateSumCorrectly()
        {
            var exp1 = new IntegerExpression("1");
            exp1.Operator = Operator.Plus;
            var exp2 = new IntegerExpression("2");
            exp1.Next = exp2;

            var result = exp1.MergeWithNext();

            Assert.That(3d, Is.EqualTo(result.Compile().Result));
        }

        [Test]
        public void MergeWithNextWithPlusOperatorShouldSetNextPointer()
        {
            var exp1 = new IntegerExpression("1");
            exp1.Operator = Operator.Plus;
            var exp2 = new IntegerExpression("2");
            exp1.Next = exp2;

            var result = exp1.MergeWithNext();

            Assert.That(result.Next, Is.Null);
        }

        //[Test]
        //public void CompileShouldHandlePercent()
        //{
        //    var exp1 = new IntegerExpression("1");
        //    exp1.Operator = Operator.Percent;
        //    exp1.Next = ConstantExpressions.Percent;
        //    var result = exp1.Compile();
        //    Assert.That(0.01, Is.EqualTo(result.Result));
        //    Assert.That(DataType.Decimal, Is.EqualTo(result.DataType));
        //}
    }
}
