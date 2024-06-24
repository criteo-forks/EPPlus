using System;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml.FormulaParsing.Excel.Functions;
using OfficeOpenXml.FormulaParsing.Excel;

namespace EPPlusTest.Excel.Functions
{
    [TestFixture]
    public class FunctionArgumentTests
    {
        [Test]
        public void ShouldSetExcelState()
        {
            var arg = new FunctionArgument(2);
            arg.SetExcelStateFlag(ExcelCellState.HiddenCell);
            Assert.That(arg.ExcelStateFlagIsSet(ExcelCellState.HiddenCell));
        }

        [Test]
        public void ExcelStateFlagIsSetShouldReturnFalseWhenNotSet()
        {
            var arg = new FunctionArgument(2);
            Assert.That(!arg.ExcelStateFlagIsSet(ExcelCellState.HiddenCell));
        }
    }
}
