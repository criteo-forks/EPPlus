using System;
using System.IO;
using NUnit.Framework;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.IntegrationTests
{
    [TestFixture]
    public class CalcExtensionsTests
    {
        [Test]
        public void ShouldCalculateChainTest()
        {
            var package = new ExcelPackage(new FileInfo("c:\\temp\\chaintest.xlsx"));
            package.Workbook.Calculate();
        }

        [Test]
        public void CalculateTest()
        {
            //var pck = new ExcelPackage();
            //var ws = pck.Workbook.Worksheets.Add("Calc1");

            //ws.SetValue("A1", (short)1);
            //var v = pck.Workbook.FormulaParserManager.Parse("2.5-Calc1!A1+abs(3.0)-SIN(3)");
            //Assert.Equals(4.358879992, Math.Round((double)v, 9));

            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");

            ws.SetValue("A1", (short)1);
            var v = pck.Workbook.FormulaParserManager.Parse("2.5-Calc1!A1+ABS(-3.0)-SIN(3)*abs(5)");
            Assert.Equals(3.79439996, Math.Round((double)v,9));
        }

        [Test]
        public void CalculateTest2()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");

            ws.SetValue("A1", (short)1);
            var v = pck.Workbook.FormulaParserManager.Parse("3*(2+5.5*2)+2*0.5+3");
            Assert.Equals(43, Math.Round((double)v, 9));
        }
    }
}
