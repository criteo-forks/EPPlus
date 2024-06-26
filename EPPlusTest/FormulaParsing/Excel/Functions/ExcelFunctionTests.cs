﻿using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using EPPlusTest.FormulaParsing.TestHelpers;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest.Excel.Functions
{
    [TestFixture]
    public class ExcelFunctionTests
    {
        private class ExcelFunctionTester : ExcelFunction
        {
            public IEnumerable<ExcelDoubleCellValue> ArgsToDoubleEnumerableImpl(IEnumerable<FunctionArgument> args)
            {
                return ArgsToDoubleEnumerable(args, ParsingContext.Create());
            }
            #region Other members
            public override CompileResult Execute(IEnumerable<FunctionArgument> arguments, ParsingContext context)
            {
                throw new NotImplementedException();
            }
            #endregion
        }

        [Test]
        public void ArgsToDoubleEnumerableShouldHandleInnerEnumerables()
        {
            var args = FunctionsHelper.CreateArgs(1, 2, FunctionsHelper.CreateArgs(3, 4));
            var tester = new ExcelFunctionTester();
            var result = tester.ArgsToDoubleEnumerableImpl(args);
            Assert.That(4, Is.EqualTo(result.Count()));
        }
    }
}
