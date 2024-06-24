using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using ExpGraph = OfficeOpenXml.FormulaParsing.ExpressionGraph.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.Excel.Operators;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestFixture]
    public class ExpressionCompilerTests
    {
        private IExpressionCompiler _expressionCompiler;
        private ExpGraph _graph;
        
        [SetUp]
        public void Setup()
        {
            _expressionCompiler = new ExpressionCompiler();
            _graph = new ExpGraph();
        }

        [Test]
        public void ShouldCompileTwoInterExpressionsToCorrectResult()
        {
            var exp1 = new IntegerExpression("2");
            exp1.Operator = Operator.Plus;
            _graph.Add(exp1);
            var exp2 = new IntegerExpression("2");
            _graph.Add(exp2);

            var result = _expressionCompiler.Compile(_graph.Expressions);

            Assert.That(4d, Is.EqualTo(result.Result));
        }


        [Test]
        public void CompileShouldMultiplyGroupExpressionWithFollowingIntegerExpression()
        {
            var groupExpression = new GroupExpression(false);
            groupExpression.AddChild(new IntegerExpression("2"));
            groupExpression.Children.First().Operator = Operator.Plus;
            groupExpression.AddChild(new IntegerExpression("3"));
            groupExpression.Operator = Operator.Multiply;

            _graph.Add(groupExpression);
            _graph.Add(new IntegerExpression("2"));

            var result = _expressionCompiler.Compile(_graph.Expressions);

            Assert.That(10d, Is.EqualTo(result.Result));
        }

        [Test]
        public void CompileShouldCalculateMultipleExpressionsAccordingToPrecedence()
        {
            var exp1 = new IntegerExpression("2");
            exp1.Operator = Operator.Multiply;
            _graph.Add(exp1);
            var exp2 = new IntegerExpression("2");
            exp2.Operator = Operator.Plus;
            _graph.Add(exp2);
            var exp3 = new IntegerExpression("2");
            exp3.Operator = Operator.Multiply;
            _graph.Add(exp3);
            var exp4 = new IntegerExpression("2");
            _graph.Add(exp4);

            var result = _expressionCompiler.Compile(_graph.Expressions);

            Assert.That(8d, Is.EqualTo(result.Result));
        }
    }
}
