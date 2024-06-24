using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.CompileStrategy;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing.Excel.Operators;
using OfficeOpenXml.FormulaParsing;
using FakeItEasy;

namespace EPPlusTest.FormulaParsing.ExpressionGraph
{
    [TestFixture]
    public class ExpressionGraphBuilderTests
    {
        private IExpressionGraphBuilder _graphBuilder;
        private ExcelDataProvider _excelDataProvider;

        [SetUp]
        public void Setup()
        {
            _excelDataProvider = A.Fake<ExcelDataProvider>();
            var parsingContext = ParsingContext.Create();
            _graphBuilder = new ExpressionGraphBuilder(_excelDataProvider, parsingContext);
        }

        [TearDown]
        public void Cleanup()
        {

        }

        [Test]
        public void BuildShouldNotUseStringIdentifyersWhenBuildingStringExpression()
        {
            var tokens = new List<Token>
            {
                new Token("'", TokenType.String),
                new Token("abc", TokenType.StringContent),
                new Token("'", TokenType.String)
            };

            var result = _graphBuilder.Build(tokens);

            Assert.That(1, Is.EqualTo(result.Expressions.Count()));
        }

        [Test]
        public void BuildShouldNotEvaluateExpressionsWithinAString()
        {
            var tokens = new List<Token>
            {
                new Token("'", TokenType.String),
                new Token("1 + 2", TokenType.StringContent),
                new Token("'", TokenType.String)
            };

            var result = _graphBuilder.Build(tokens);

            Assert.That("1 + 2", Is.EqualTo(result.Expressions.First().Compile().Result));
        }

        [Test]
        public void BuildShouldSetOperatorOnGroupExpressionCorrectly()
        {
            var tokens = new List<Token>
            {
                new Token("(", TokenType.OpeningParenthesis),
                new Token("2", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("4", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis),
                new Token("*", TokenType.Operator),
                new Token("2", TokenType.Integer)
            };
            var result = _graphBuilder.Build(tokens);

            Assert.That(Operator.Multiply.Operator, Is.EqualTo(result.Expressions.First().Operator.Operator));

        }

        [Test]
        public void BuildShouldSetChildrenOnGroupExpression()
        {
            var tokens = new List<Token>
            {
                new Token("(", TokenType.OpeningParenthesis),
                new Token("2", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("4", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis),
                new Token("*", TokenType.Operator),
                new Token("2", TokenType.Integer)
            };
            var result = _graphBuilder.Build(tokens);

            Assert.That(result.Expressions.First(), Is.InstanceOf<GroupExpression>());
            Assert.That(2, Is.EqualTo(result.Expressions.First().Children.Count()));
        }

        [Test]
        public void BuildShouldSetNextOnGroupedExpression()
        {
            var tokens = new List<Token>
            {
                new Token("(", TokenType.OpeningParenthesis),
                new Token("2", TokenType.Integer),
                new Token("+", TokenType.Operator),
                new Token("4", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis),
                new Token("*", TokenType.Operator),
                new Token("2", TokenType.Integer)
            };
            var result = _graphBuilder.Build(tokens);

            Assert.That(result.Expressions.First().Next, Is.Not.Null);
            Assert.That(result.Expressions.First().Next, Is.InstanceOf<IntegerExpression>());

        }

        [Test]
        public void BuildShouldBuildFunctionExpressionIfFirstTokenIsFunction()
        {
            var tokens = new List<Token>
            {
                new Token("CStr", TokenType.Function),
                new Token("(", TokenType.OpeningParenthesis),
                new Token("2", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis),
            };
            var result = _graphBuilder.Build(tokens);

            Assert.That(1, Is.EqualTo(result.Expressions.Count()));
            Assert.That(result.Expressions.First(), Is.InstanceOf<FunctionExpression>());
        }

        [Test]
        public void BuildShouldSetChildrenOnFunctionExpression()
        {
            var tokens = new List<Token>
            {
                new Token("CStr", TokenType.Function),
                new Token("(", TokenType.OpeningParenthesis),
                new Token("2", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis)
            };
            var result = _graphBuilder.Build(tokens);

            Assert.That(1, Is.EqualTo(result.Expressions.First().Children.Count()));
            Assert.That(result.Expressions.First().Children.First(), Is.InstanceOf<GroupExpression>());
            Assert.That(result.Expressions.First().Children.First().Children.First(), Is.InstanceOf<IntegerExpression>());
            Assert.That(2d, Is.EqualTo(result.Expressions.First().Children.First().Compile().Result));
        }

        [Test]
        public void BuildShouldAddOperatorToFunctionExpression()
        {
            var tokens = new List<Token>
            {
                new Token("CStr", TokenType.Function),
                new Token("(", TokenType.OpeningParenthesis),
                new Token("2", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis),
                new Token("&", TokenType.Operator),
                new Token("A", TokenType.StringContent)
            };
            var result = _graphBuilder.Build(tokens);

            Assert.That(1, Is.EqualTo(result.Expressions.First().Children.Count()));
            Assert.That(2, Is.EqualTo(result.Expressions.Count()));
        }

        [Test]
        public void BuildShouldAddCommaSeparatedFunctionArgumentsAsChildrenToFunctionExpression()
        {
            var tokens = new List<Token>
            {
                new Token("Text", TokenType.Function),
                new Token("(", TokenType.OpeningParenthesis),
                new Token("2", TokenType.Integer),
                new Token(",", TokenType.Comma),
                new Token("3", TokenType.Integer),
                new Token(")", TokenType.ClosingParenthesis),
                new Token("&", TokenType.Operator),
                new Token("A", TokenType.StringContent)
            };

            var result = _graphBuilder.Build(tokens);

            Assert.That(2, Is.EqualTo(result.Expressions.First().Children.Count()));
        }

        [Test]
        public void BuildShouldCreateASingleExpressionOutOfANegatorAndANumericToken()
        {
            var tokens = new List<Token>
            {
                new Token("-", TokenType.Negator),
                new Token("2", TokenType.Integer),
            };

            var result = _graphBuilder.Build(tokens);

            Assert.That(1, Is.EqualTo(result.Expressions.Count()));
            Assert.That(-2d, Is.EqualTo(result.Expressions.First().Compile().Result));
        }

        [Test]
        public void BuildShouldHandleEnumerableTokens()
        {
            var tokens = new List<Token>
            {
                new Token("Text", TokenType.Function),
                new Token("(", TokenType.OpeningParenthesis),
                new Token("{", TokenType.OpeningEnumerable),
                new Token("2", TokenType.Integer),
                new Token(",", TokenType.Comma),
                new Token("3", TokenType.Integer),
                new Token("}", TokenType.ClosingEnumerable),
                new Token(")", TokenType.ClosingParenthesis)
            };

            var result = _graphBuilder.Build(tokens);
            var funcArgExpression = result.Expressions.First().Children.First();
            Assert.That(funcArgExpression, Is.InstanceOf<FunctionArgumentExpression>());

            var enumerableExpression = funcArgExpression.Children.First();

            Assert.That(enumerableExpression, Is.InstanceOf<EnumerableExpression>());
            Assert.That(2, Is.EqualTo(enumerableExpression.Children.Count()), "Enumerable.Count was not 2");
        }

        [Test]
        public void ShouldHandleInnerFunctionCall2()
        {
            var ctx = ParsingContext.Create();
            const string formula = "IF(3>2;\"Yes\";\"No\")";
            var tokenizer = new SourceCodeTokenizer(ctx.Configuration.FunctionRepository, ctx.NameValueProvider);
            var tokens = tokenizer.Tokenize(formula);
            var expression = _graphBuilder.Build(tokens);
            Assert.That(1, Is.EqualTo(expression.Expressions.Count()));

            var compiler = new ExpressionCompiler(new ExpressionConverter(), new CompileStrategyFactory());
            var result = compiler.Compile(expression.Expressions);
            Assert.That("Yes", Is.EqualTo(result.Result));
        }

        [Test]
        public void ShouldHandleInnerFunctionCall3()
        {
            var ctx = ParsingContext.Create();
            const string formula = "IF(I10>=0;IF(O10>I10;((O10-I10)*$B10)/$C$27;IF(O10<0;(O10*$B10)/$C$27;\"\"));IF(O10<0;((O10-I10)*$B10)/$C$27;IF(O10>0;(O10*$B10)/$C$27;)))";
            var tokenizer = new SourceCodeTokenizer(ctx.Configuration.FunctionRepository, ctx.NameValueProvider);
            var tokens = tokenizer.Tokenize(formula);
            var expression = _graphBuilder.Build(tokens);
            Assert.That(1, Is.EqualTo(expression.Expressions.Count()));
            var exp1 = expression.Expressions.First();
            Assert.That(3, Is.EqualTo(exp1.Children.Count()));
        }
        [Test]
        public void RemoveDuplicateOperators1()
        {
            var ctx = ParsingContext.Create();
            const string formula = "++1--2++-3+-1----3-+2";
            var tokenizer = new SourceCodeTokenizer(ctx.Configuration.FunctionRepository, ctx.NameValueProvider);
            var tokens = tokenizer.Tokenize(formula).ToList();
            var expression = _graphBuilder.Build(tokens);
            Assert.That(11, Is.EqualTo(tokens.Count()));
            Assert.That("+", Is.EqualTo(tokens[1].Value));
            Assert.That("-", Is.EqualTo(tokens[3].Value));
            Assert.That("-", Is.EqualTo(tokens[5].Value));
            Assert.That("+", Is.EqualTo(tokens[7].Value));
            Assert.That("-", Is.EqualTo(tokens[9].Value));
        }
        [Test]
        public void RemoveDuplicateOperators2()
        {
            var ctx = ParsingContext.Create();
            const string formula = "++-1--(---2)++-3+-1----3-+2";
            var tokenizer = new SourceCodeTokenizer(ctx.Configuration.FunctionRepository, ctx.NameValueProvider);
            var tokens = tokenizer.Tokenize(formula).ToList();
        }

        [Test]
        public void BuildExcelAddressExpressionSimple()
        {
            var tokens = new List<Token>
            {
                new Token("A1", TokenType.ExcelAddress)
            };

            var result = _graphBuilder.Build(tokens);
            Assert.That(result.Expressions.First(), Is.InstanceOf<ExcelAddressExpression>());
        }
    }
}
