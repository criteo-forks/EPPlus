using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing.LexicalAnalysis;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.FormulaParsing.LexicalAnalysis
{
    [TestFixture]
    public class SourceCodeTokenizerTests
    {
        private SourceCodeTokenizer _tokenizer;

        [SetUp]
        public void Setup()
        {
            var context = ParsingContext.Create();
            _tokenizer = new SourceCodeTokenizer(context.Configuration.FunctionRepository, null);
        }

        [TearDown]
        public void Cleanup()
        {
        }

        [Test]
        public void ShouldCreateTokensForStringCorrectly()
        {
            var input = "\"abc123\"";
            var tokens = _tokenizer.Tokenize(input);

            Assert.That(3, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.String, Is.EqualTo(tokens.First().TokenType));
            Assert.That(TokenType.StringContent, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That(TokenType.String, Is.EqualTo(tokens.Last().TokenType));
        }

        [Test]
        public void ShouldTokenizeStringCorrectly()
        {
            var input = "\"ab(c)d\"";
            var tokens = _tokenizer.Tokenize(input);

            Assert.That(3, Is.EqualTo(tokens.Count()));
        }

        [Test]
        public void ShouldHandleWhitespaceCorrectly()
        {
            var input = @"""          """;
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(3, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.StringContent, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That(10, Is.EqualTo(tokens.ElementAt(1).Value.Length));
        }

        [Test]
        public void ShouldCreateTokensForFunctionCorrectly()
        {
            var input = "Text(2)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.That(4, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Function, Is.EqualTo(tokens.First().TokenType));
            Assert.That(TokenType.OpeningParenthesis, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(2).TokenType));
            Assert.That("2", Is.EqualTo(tokens.ElementAt(2).Value));
            Assert.That(TokenType.ClosingParenthesis, Is.EqualTo(tokens.Last().TokenType));
        }

        [Test]
        public void ShouldHandleMultipleCharOperatorCorrectly()
        {
            var input = "1 <= 2";
            var tokens = _tokenizer.Tokenize(input);

            Assert.That(3, Is.EqualTo(tokens.Count()));
            Assert.That("<=", Is.EqualTo(tokens.ElementAt(1).Value));
            Assert.That(TokenType.Operator, Is.EqualTo(tokens.ElementAt(1).TokenType));
        }

        [Test]
        public void ShouldCreateTokensForEnumerableCorrectly()
        {
            var input = "Text({1;2})";
            var tokens = _tokenizer.Tokenize(input);

            Assert.That(8, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.OpeningEnumerable, Is.EqualTo(tokens.ElementAt(2).TokenType));
            Assert.That(TokenType.ClosingEnumerable, Is.EqualTo(tokens.ElementAt(6).TokenType));
        }

        [Test]
        public void ShouldCreateTokensForExcelAddressCorrectly()
        {
            var input = "Text(A1)";
            var tokens = _tokenizer.Tokenize(input);

            Assert.That(TokenType.ExcelAddress, Is.EqualTo(tokens.ElementAt(2).TokenType));
        }

        [Test]
        public void ShouldCreateTokenForPercentAfterDecimal()
        {
            var input = "1,23%";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(TokenType.Percent, Is.EqualTo(tokens.Last().TokenType));
        }

        [Test]
        public void ShouldIgnoreTwoSubsequentStringIdentifyers()
        {
            var input = "\"hello\"\"world\"";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(3, Is.EqualTo(tokens.Count()));
            Assert.That("hello\"world", Is.EqualTo(tokens.ElementAt(1).Value));
        }

        [Test]
        public void ShouldIgnoreTwoSubsequentStringIdentifyers2()
        {
            //using (var pck = new ExcelPackage(new FileInfo("c:\\temp\\QuoteIssue2.xlsx")))
            //{
            //    pck.Workbook.Worksheets.First().Calculate();
            //}
            var input = "\"\"\"\"\"\"";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(TokenType.StringContent, Is.EqualTo(tokens.ElementAt(1).TokenType));
        }

        [Test]
        public void TokenizerShouldIgnoreOperatorInString()
        {
            var input = "\"*\"";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(TokenType.StringContent, Is.EqualTo(tokens.ElementAt(1).TokenType));
        }

        [Test]
        public void TokenizerShouldHandleWorksheetNameWithMinus()
        {
            var input = "'A-B'!A1";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(1, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.ExcelAddress, Is.EqualTo(tokens.ElementAt(0).TokenType));
        }

        [Test]
        public void TestBug9_12_14()
        {
            //(( W60 -(- W63 )-( W29 + W30 + W31 ))/( W23 + W28 + W42 - W51 )* W4 )
            using (var pck = new ExcelPackage())
            {
                var ws1 = pck.Workbook.Worksheets.Add("test");
                for (var x = 1; x <= 10; x++)
                {
                    ws1.Cells[x, 1].Value = x;
                }

                ws1.Cells["A11"].Formula = "(( A1 -(- A2 )-( A3 + A4 + A5 ))/( A6 + A7 + A8 - A9 )* A5 )";
                //ws1.Cells["A11"].Formula = "(-A2 + 1 )";
                ws1.Calculate();
                var result = ws1.Cells["A11"].Value;
                Assert.That(-3.75, Is.EqualTo(result));
            }
        }

        [Test]
        public void TokenizeStripsLeadingPlusSign()
        {
            var input = @"+3-3";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(3, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(0).TokenType));
            Assert.That(TokenType.Operator, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(2).TokenType));
        }

        [Test]
        public void TokenizeStripsLeadingDoubleNegator()
        {
            var input = @"--3-3";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(3, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(0).TokenType));
            Assert.That(TokenType.Operator, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(2).TokenType));
        }

        [Test]
        public void TokenizeHandlesPositiveNegator()
        {
            var input = @"+-3-3";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(4, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Negator, Is.EqualTo(tokens.ElementAt(0).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That(TokenType.Operator, Is.EqualTo(tokens.ElementAt(2).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(3).TokenType));
        }

        [Test]
        public void TokenizeHandlesNegatorPositive()
        {
            var input = @"-+3-3";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(4, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Negator, Is.EqualTo(tokens.ElementAt(0).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That(TokenType.Operator, Is.EqualTo(tokens.ElementAt(2).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(3).TokenType));
        }

        [Test]
        public void TokenizeStripsLeadingPlusSignFromFirstFunctionArgument()
        {
            var input = @"SUM(+3-3,5)";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(8, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Function, Is.EqualTo(tokens.ElementAt(0).TokenType));
            Assert.That(TokenType.OpeningParenthesis, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(2).TokenType));
            Assert.That(TokenType.Operator, Is.EqualTo(tokens.ElementAt(3).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(4).TokenType));
            Assert.That(TokenType.Comma, Is.EqualTo(tokens.ElementAt(5).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(6).TokenType));
            Assert.That(TokenType.ClosingParenthesis, Is.EqualTo(tokens.ElementAt(7).TokenType));
        }

        [Test]
        public void TokenizeStripsLeadingPlusSignFromSecondFunctionArgument()
        {
            var input = @"SUM(5,+3-3)";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(8, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Function, Is.EqualTo(tokens.ElementAt(0).TokenType));
            Assert.That(TokenType.OpeningParenthesis, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(2).TokenType));
            Assert.That(TokenType.Comma, Is.EqualTo(tokens.ElementAt(3).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(4).TokenType));
            Assert.That(TokenType.Operator, Is.EqualTo(tokens.ElementAt(5).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(6).TokenType));
            Assert.That(TokenType.ClosingParenthesis, Is.EqualTo(tokens.ElementAt(7).TokenType));
        }

        [Test]
        public void TokenizeStripsLeadingDoubleNegatorFromFirstFunctionArgument()
        {
            var input = @"SUM(--3-3,5)";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(8, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Function, Is.EqualTo(tokens.ElementAt(0).TokenType));
            Assert.That(TokenType.OpeningParenthesis, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(2).TokenType));
            Assert.That(TokenType.Operator, Is.EqualTo(tokens.ElementAt(3).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(4).TokenType));
            Assert.That(TokenType.Comma, Is.EqualTo(tokens.ElementAt(5).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(6).TokenType));
            Assert.That(TokenType.ClosingParenthesis, Is.EqualTo(tokens.ElementAt(7).TokenType));
        }

        [Test]
        public void TokenizeStripsLeadingDoubleNegatorFromSecondFunctionArgument()
        {
            var input = @"SUM(5,--3-3)";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(8, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Function, Is.EqualTo(tokens.ElementAt(0).TokenType));
            Assert.That(TokenType.OpeningParenthesis, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(2).TokenType));
            Assert.That(TokenType.Comma, Is.EqualTo(tokens.ElementAt(3).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(4).TokenType));
            Assert.That(TokenType.Operator, Is.EqualTo(tokens.ElementAt(5).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(6).TokenType));
            Assert.That(TokenType.ClosingParenthesis, Is.EqualTo(tokens.ElementAt(7).TokenType));
        }

        [Test]
        public void TokenizeHandlesPositiveNegatorAsFirstFunctionArgument()
        {
            var input = @"SUM(+-3-3,5)";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(9, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Function, Is.EqualTo(tokens.ElementAt(0).TokenType));
            Assert.That(TokenType.OpeningParenthesis, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That(TokenType.Negator, Is.EqualTo(tokens.ElementAt(2).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(3).TokenType));
            Assert.That(TokenType.Operator, Is.EqualTo(tokens.ElementAt(4).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(5).TokenType));
            Assert.That(TokenType.Comma, Is.EqualTo(tokens.ElementAt(6).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(7).TokenType));
            Assert.That(TokenType.ClosingParenthesis, Is.EqualTo(tokens.ElementAt(8).TokenType));
        }

        [Test]
        public void TokenizeHandlesNegatorPositiveAsFirstFunctionArgument()
        {
            var input = @"SUM(-+3-3,5)";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(9, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Function, Is.EqualTo(tokens.ElementAt(0).TokenType));
            Assert.That(TokenType.OpeningParenthesis, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That(TokenType.Negator, Is.EqualTo(tokens.ElementAt(2).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(3).TokenType));
            Assert.That(TokenType.Operator, Is.EqualTo(tokens.ElementAt(4).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(5).TokenType));
            Assert.That(TokenType.Comma, Is.EqualTo(tokens.ElementAt(6).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(7).TokenType));
            Assert.That(TokenType.ClosingParenthesis, Is.EqualTo(tokens.ElementAt(8).TokenType));
        }

        [Test]
        public void TokenizeHandlesPositiveNegatorAsSecondFunctionArgument()
        {
            var input = @"SUM(5,+-3-3)";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(9, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Function, Is.EqualTo(tokens.ElementAt(0).TokenType));
            Assert.That(TokenType.OpeningParenthesis, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(2).TokenType));
            Assert.That(TokenType.Comma, Is.EqualTo(tokens.ElementAt(3).TokenType));
            Assert.That(TokenType.Negator, Is.EqualTo(tokens.ElementAt(4).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(5).TokenType));
            Assert.That(TokenType.Operator, Is.EqualTo(tokens.ElementAt(6).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(7).TokenType));
            Assert.That(TokenType.ClosingParenthesis, Is.EqualTo(tokens.ElementAt(8).TokenType));
        }

        [Test]
        public void TokenizeHandlesNegatorPositiveAsSecondFunctionArgument()
        {
            var input = @"SUM(5,-+3-3)";
            var tokens = _tokenizer.Tokenize(input);
            Assert.That(9, Is.EqualTo(tokens.Count()));
            Assert.That(TokenType.Function, Is.EqualTo(tokens.ElementAt(0).TokenType));
            Assert.That(TokenType.OpeningParenthesis, Is.EqualTo(tokens.ElementAt(1).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(2).TokenType));
            Assert.That(TokenType.Comma, Is.EqualTo(tokens.ElementAt(3).TokenType));
            Assert.That(TokenType.Negator, Is.EqualTo(tokens.ElementAt(4).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(5).TokenType));
            Assert.That(TokenType.Operator, Is.EqualTo(tokens.ElementAt(6).TokenType));
            Assert.That(TokenType.Integer, Is.EqualTo(tokens.ElementAt(7).TokenType));
            Assert.That(TokenType.ClosingParenthesis, Is.EqualTo(tokens.ElementAt(8).TokenType));
        }
    }
}
