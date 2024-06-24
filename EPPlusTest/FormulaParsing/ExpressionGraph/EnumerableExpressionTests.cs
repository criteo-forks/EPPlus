using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestFixture]
    public class EnumerableExpressionTests
    {
        [Test]
        public void CompileShouldReturnEnumerableOfCompiledChildExpressions()
        {
            var expression = new EnumerableExpression();
            expression.AddChild(new IntegerExpression("2"));
            expression.AddChild(new IntegerExpression("3"));
            var result = expression.Compile();

            Assert.That(result.Result, Is.InstanceOf<IEnumerable<object>>());
            var resultList = (IEnumerable<object>)result.Result;
            Assert.That(2d, Is.EqualTo(resultList.ElementAt(0)));
            Assert.That(3d, Is.EqualTo(resultList.ElementAt(1)));
        }
    }
}
