using System;
using System.Text;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Logical;
using OfficeOpenXml;

namespace EPPlusTest.Excel.Functions
{
    [TestFixture]
    public class LogicalFunctionsTests
    {
        private ParsingContext _parsingContext = ParsingContext.Create();

        [Test]
        public void IfShouldReturnCorrectResult()
        {
            var func = new If();
            var args = FunctionsHelper.CreateArgs(true, "A", "B");
            var result = func.Execute(args, _parsingContext);
            Assert.That("A", Is.EqualTo(result.Result));
        }

        [Test] [Explicit]
        public void IfShouldIgnoreCase()
        {
            using (var pck = new ExcelPackage(new FileInfo(@"c:\temp\book1.xlsx")))
            {
                pck.Workbook.Calculate();
                Assert.That("Sant", Is.EqualTo(pck.Workbook.Worksheets.First().Cells["C3"].Value));
            }
        }

        [Test]
        public void NotShouldReturnFalseIfArgumentIsTrue()
        {
            var func = new Not();
            var args = FunctionsHelper.CreateArgs(true);
            var result = func.Execute(args, _parsingContext);
            Assert.That(!(bool)result.Result);
        }

        [Test]
        public void NotShouldReturnTrueIfArgumentIs0()
        {
            var func = new Not();
            var args = FunctionsHelper.CreateArgs(0);
            var result = func.Execute(args, _parsingContext);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void NotShouldReturnFalseIfArgumentIs1()
        {
            var func = new Not();
            var args = FunctionsHelper.CreateArgs(1);
            var result = func.Execute(args, _parsingContext);
            Assert.That(!(bool)result.Result);
        }

        [Test]
        public void NotShouldHandleExcelReference()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = false;
                sheet.Cells["A2"].Formula = "NOT(A1)";
                sheet.Calculate();
                Assert.That((bool)sheet.Cells["A2"].Value);
            }
        }

        [Test]
        public void NotShouldHandleExcelReferenceToStringFalse()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = "false";
                sheet.Cells["A2"].Formula = "NOT(A1)";
                sheet.Calculate();
                Assert.That((bool)sheet.Cells["A2"].Value);
            }
        }

        [Test]
        public void NotShouldHandleExcelReferenceToStringTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = "TRUE";
                sheet.Cells["A2"].Formula = "NOT(A1)";
                sheet.Calculate();
                Assert.That(!(bool)sheet.Cells["A2"].Value);
            }
        }

        [Test]
        public void AndShouldHandleStringLiteralTrue()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("sheet1");
                sheet.Cells["A1"].Value = "tRuE";
                sheet.Cells["A2"].Formula = "AND(\"TRUE\", A1)";
                sheet.Calculate();
                Assert.That((bool)sheet.Cells["A2"].Value);
            }
        }

        [Test]
        public void AndShouldReturnTrueIfAllArgumentsAreTrue()
        {
            var func = new And();
            var args = FunctionsHelper.CreateArgs(true, true, true);
            var result = func.Execute(args, _parsingContext);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void AndShouldReturnTrueIfAllArgumentsAreTrueOr1()
        {
            var func = new And();
            var args = FunctionsHelper.CreateArgs(true, true, 1, true, 1);
            var result = func.Execute(args, _parsingContext);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void AndShouldReturnFalseIfOneArgumentIsFalse()
        {
            var func = new And();
            var args = FunctionsHelper.CreateArgs(true, false, true);
            var result = func.Execute(args, _parsingContext);
            Assert.That(!(bool)result.Result);
        }

        [Test]
        public void AndShouldReturnFalseIfOneArgumentIs0()
        {
            var func = new And();
            var args = FunctionsHelper.CreateArgs(true, 0, true);
            var result = func.Execute(args, _parsingContext);
            Assert.That(!(bool)result.Result);
        }

        [Test]
        public void OrShouldReturnTrueIfOneArgumentIsTrue()
        {
            var func = new Or();
            var args = FunctionsHelper.CreateArgs(true, false, false);
            var result = func.Execute(args, _parsingContext);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void OrShouldReturnTrueIfOneArgumentIsTrueString()
        {
            var func = new Or();
            var args = FunctionsHelper.CreateArgs("true", "FALSE", false);
            var result = func.Execute(args, _parsingContext);
            Assert.That((bool)result.Result);
        }

        [Test]
        public void IfErrorShouldReturnSecondArgIfCriteriaEvaluatesAsAnError()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "IFERROR(0/0, \"hello\")";
                s1.Calculate();
                Assert.That("hello", Is.EqualTo(s1.Cells["A1"].Value));
            }
        }

        [Test]
        public void IfErrorShouldReturnSecondArgIfCriteriaEvaluatesAsAnError2()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "IFERROR(A2, \"hello\")";
                s1.Cells["A2"].Formula = "23/0";
                s1.Calculate();
                Assert.That("hello", Is.EqualTo(s1.Cells["A1"].Value));
            }
        }

        [Test]
        public void IfErrorShouldReturnResultOfFormulaIfNoError()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "IFERROR(A2, \"hello\")";
                s1.Cells["A2"].Value = "hi there";
                s1.Calculate();
                Assert.That("hi there", Is.EqualTo(s1.Cells["A1"].Value));
            }
        }

        [Test]
        public void IfNaShouldReturnSecondArgIfCriteriaEvaluatesAsAnError2()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "IFERROR(A2, \"hello\")";
                s1.Cells["A2"].Value = ExcelErrorValue.Create(eErrorType.NA);
                s1.Calculate();
                Assert.That("hello", Is.EqualTo(s1.Cells["A1"].Value));
            }
        }

        [Test]
        public void IfNaShouldReturnResultOfFormulaIfNoError()
        {
            using (var package = new ExcelPackage())
            {
                var s1 = package.Workbook.Worksheets.Add("test");
                s1.Cells["A1"].Formula = "IFNA(A2, \"hello\")";
                s1.Cells["A2"].Value = "hi there";
                s1.Calculate();
                Assert.That("hi there", Is.EqualTo(s1.Cells["A1"].Value));
            }
        }
    }
}
