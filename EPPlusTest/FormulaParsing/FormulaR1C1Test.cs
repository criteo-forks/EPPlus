using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing
{
    [TestFixture]
    public class FormulaR1C1Tests
    {
        private ExcelPackage _package;
        private ExcelWorksheet _sheet;
        private ExcelWorksheet _sheet2;
        [SetUp]
        public void Initialize()
        {
            _package = new ExcelPackage();
            var s1 = _package.Workbook.Worksheets.Add("test");
            s1.Cells["A1"].Value = 1;
            s1.Cells["A2"].Value = 2;
            s1.Cells["A3"].Value = 3;
            s1.Cells["A4"].Value = 4;

            s1.Cells["B1"].Value = 5;
            s1.Cells["B2"].Value = 6;
            s1.Cells["B3"].Value = 7;
            s1.Cells["B4"].Value = 8;

            s1.Cells["C1"].Value = 5;
            s1.Cells["C2"].Value = 6;
            s1.Cells["C3"].Value = 7;
            s1.Cells["C4"].Value = 8;

            _sheet = s1;
            _sheet2 = _package.Workbook.Worksheets.Add("test2",s1);
        }

        [TearDown]
        public void Cleanup()
        {
            _package.Dispose();
        }

        [Test]
        public void RC2()
        {
            string fR1C1 = "RC2";
            _sheet.Cells[5, 1].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 1].Formula;
            _sheet.Cells[5, 1].Formula = f;
            Assert.Equals(fR1C1, _sheet.Cells[5,1].FormulaR1C1);
        }
        [Test]
        public void C()
        {
            string fR1C1 = "SUMIFS(C,C2,RC1)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            _sheet.Cells[5, 3].Formula = f;
            Assert.Equals(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
        }
        [Test]
        public void C2Abs()
        {
            string fR1C1 = "SUM(C2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.That("SUM($B:$B)", Is.EqualTo(f));
        }
        [Test]
        public void C2AbsWithSheet()
        {
            string fR1C1 = "SUM(A!C2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.That("SUM(A!$B:$B)", Is.EqualTo(f));
        }
        [Test]
        public void C2()
        {
            string fR1C1 = "SUM(C2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            _sheet.Cells[5, 3].Formula = f;
            Assert.Equals(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
        }
        [Test]
        public void R2Abs()
        {
            string fR1C1 = "SUM(R2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.Equals("SUM($2:$2)",f);

            fR1C1 = "SUM(TEST2!R2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            f = _sheet.Cells[5, 3].Formula;
            Assert.That("SUM(TEST2!$2:$2)", Is.EqualTo(f));

        }
        [Test]
        public void R2()
        {
            string fR1C1 = "SUM(R2)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            _sheet.Cells[5, 3].Formula = f;
            Assert.Equals(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
        }
        [Test]
        public void RCRelativeToAB()
        {
            string fR1C1 = "SUMIFS(C,C2,RC1)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.Equals("SUMIFS(C:C,$B:$B,$A5)", f);
        }
        [Test]
        public void RRelativeToAB()
        {
            string fR1C1 = "SUMIFS(R,C2,RC1)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            Assert.Equals("SUMIFS(5:5,$B:$B,$A5)", f);
        }
        [Test]
        public void RCRelativeToABToR1C1()
        {
            string fR1C1 = "SUMIFS(C,C2,RC1)";
            _sheet.Cells[5, 3].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 3].Formula;
            _sheet.Cells[5, 3].Formula = f;
            Assert.Equals(fR1C1, _sheet.Cells[5, 3].FormulaR1C1);
        }
        [Test]
        public void RCRelativeToABToR1C1_2()
        {
            string fR1C1 = "SUM(RC9:RC[-1])";
            _sheet.Cells[5, 13].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[5, 13].Formula;
            Assert.That("SUM($I5:L5)", Is.EqualTo(f));
            _sheet.Cells[5, 13].Formula = f;
            Assert.Equals(fR1C1, _sheet.Cells[5, 13].FormulaR1C1);

            //"RC{colShort} - SUM(RC21:RC12)";
        }
        [Test]
        public void RCFixToABToR1C1_2()
        {
            string fR1C1 = "RC28-SUM(RC12:RC21)";
            _sheet.Cells[6, 13].FormulaR1C1 = fR1C1;
            string f = _sheet.Cells[6, 13].Formula;
            Assert.That("$AB6-SUM($L6:$U6)", Is.EqualTo(f));
            _sheet.Cells[6, 13].Formula = f;
            Assert.Equals(fR1C1, _sheet.Cells[6, 13].FormulaR1C1);
        }
        [Test]
        public void SimpleRelativeR1C1()
        {
            string fR1C1 = "R[-1]C[-5]";
            var c = _sheet.Cells[7, 7];
            c.FormulaR1C1 = fR1C1;
            string f = c.Formula;
            Assert.That("B6", Is.EqualTo(f));
            c.Formula = f;
            Assert.That(fR1C1, Is.EqualTo(c.FormulaR1C1));
        }
        [Test]
        public void SimpleAbsR1C1()
        {
            string fR1C1 = "R1C5";
            var c = _sheet.Cells[8, 8];
            c.FormulaR1C1 = fR1C1;
            string f = c.Formula;
            Assert.That("$E$1", Is.EqualTo(f));
            c.Formula = f;
            Assert.That(fR1C1, Is.EqualTo(c.FormulaR1C1));
        }
        [Test]
        public void FullTwoColumn()
        {
            string formula = "VLOOKUP(C2,A:B,1,0)";
            var c = _sheet.Cells["D2"];
            c.Formula = formula;
            Assert.AreEquals(c.FormulaR1C1, "VLOOKUP(RC[-1],C[-3]:C[-2],1,0)");
            c.FormulaR1C1 = c.FormulaR1C1;
            Assert.That(c.Formula, Is.EqualTo(formula));
        }
        [Test]
        public void FullColumn()
        {
            string formula = "VLOOKUP(C2,A:A,1,0)";
            var c = _sheet.Cells["D2"];
            c.Formula = formula;
            Assert.Equals(c.FormulaR1C1, "VLOOKUP(RC[-1],C[-3],1,0)");
            c.FormulaR1C1 = c.FormulaR1C1;
            Assert.That(c.Formula, Is.EqualTo(formula));
        }
        [Test]
        public void FullTwoRow()
        {
            string formula = "VLOOKUP(C3,1:2,1,0)";
            var c = _sheet.Cells["D3"];
            c.Formula = formula;
            Assert.Equals(c.FormulaR1C1, "VLOOKUP(RC[-1],R[-2]:R[-1],1,0)");
            c.FormulaR1C1 = c.FormulaR1C1;
            Assert.That(c.Formula, Is.EqualTo(formula));
        }
        [Test]
        public void FullRow()
        {
            string formula = "VLOOKUP(C3,1:1,1,0)";
            var c = _sheet.Cells["D3"];
            c.Formula = formula;
            Assert.Equals(c.FormulaR1C1, "VLOOKUP(RC[-1],R[-2],1,0)");
            c.FormulaR1C1 = c.FormulaR1C1;
            Assert.That(c.Formula, Is.EqualTo(formula));
        }

        [Test]
        public void OutOfRangeCol()
        {
            _sheet.Cells["a3"].FormulaR1C1 = "R[-3]C";
            Assert.That("#REF!", Is.EqualTo(_sheet.Cells["a3"].Formula));
            _sheet.Cells["a3"].FormulaR1C1 = "R[-2]C";
            Assert.That("A1", Is.EqualTo(_sheet.Cells["a3"].Formula));

            _sheet.Cells["B3"].FormulaR1C1 = "RC[-2]";
            Assert.That("#REF!", Is.EqualTo(_sheet.Cells["B3"].Formula));
            _sheet.Cells["B3"].FormulaR1C1 = "RC[-1]";
            Assert.That("A3", Is.EqualTo(_sheet.Cells["B3"].Formula));

        }
    }
}
