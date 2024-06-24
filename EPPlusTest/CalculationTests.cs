using System;
using System.Globalization;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml;
using System.IO;
using System.Diagnostics;
using OfficeOpenXml.FormulaParsing;

namespace EPPlusTest
{
    [TestFixture]
    public class CalculationTests
    {
        //[Test]
        //public void Calulation()
        //{
        //    var pck = new ExcelPackage(new FileInfo("c:\\temp\\chain.xlsx"));
        //    pck.Workbook.Calculate();
        //    Assert.That(50D, Is.EqualTo(pck.Workbook.Worksheets[1].Cells["C1"].Value));
        //}
        //[Test]
        //public void Calulation2()
        //{
        //    var pck = new ExcelPackage(new FileInfo("c:\\temp\\chainTest.xlsx"));
        //    pck.Workbook.Calculate();
        //    Assert.That(1124999960382D, Is.EqualTo(pck.Workbook.Worksheets[1].Cells["C1"].Value));
        //}
        //[Test]
        //public void Calulation3()
        //{
        //    var pck = new ExcelPackage(new FileInfo("c:\\temp\\names.xlsx"));
        //    pck.Workbook.Calculate();
        //    //Assert.That(1124999960382D, Is.EqualTo(pck.Workbook.Worksheets[1].Cells["C1"].Value));
        //}
        [Test]
        public void CalulationTestDatatypes()
        {
            var pck = new ExcelPackage();
            var ws=pck.Workbook.Worksheets.Add("Calc1");
            ws.SetValue("A1", (short)1);
            ws.SetValue("A2", (long)2);
            ws.SetValue("A3", (Single)3);
            ws.SetValue("A4", (double)4);
            ws.SetValue("A5", (Decimal)5);
            ws.SetValue("A6", (byte)6);
            ws.SetValue("A7", null);
            ws.Cells["A10"].Formula = "Sum(A1:A8)";
            ws.Cells["A11"].Formula = "SubTotal(9,A1:A8)";
            ws.Cells["A12"].Formula = "Average(A1:A8)";

            ws.Calculate();
            Assert.That(21D, Is.EqualTo(ws.Cells["a10"].Value));
            Assert.That(21D, Is.EqualTo(ws.Cells["a11"].Value));
            Assert.That(21D/6, Is.EqualTo(ws.Cells["a12"].Value));
        }
        [Test]
        public void CalculateTest()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");

            ws.SetValue("A1",( short)1);
            var v=ws.Calculate("2.5-A1+ABS(-3.0)-SIN(3)");
            Assert.Equals(4.3589, Math.Round((double)v, 4));
                        
            ws.Row(1).Hidden = true;
            v = ws.Calculate("subtotal(109,a1:a10)");
            Assert.That(0D, Is.EqualTo(v));

            v = ws.Calculate("-subtotal(9,a1:a3)");
            Assert.That(-1D, Is.EqualTo(v));
        }
        [Test]
        public void CalculateTestIsFunctions()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");

            ws.SetValue(1, 1, 1.0D);
            ws.SetFormula(1, 2, "isblank(A1:A5)");
            ws.SetFormula(1, 3, "concatenate(a1,a2,a3)");
            ws.SetFormula(1, 4, "Row()");
            ws.SetFormula(1, 5, "Row(a3)");
            ws.Calculate();
        }
        [Test] [Explicit]
        public void Calulation4()
        {
#if Core
            var dir = AppContext.BaseDirectory;
            dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
            var pck = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "FormulaTest.xlsx")));
            pck.Workbook.Calculate();
            Assert.That(490D, Is.EqualTo(pck.Workbook.Worksheets[1].Cells["D5"].Value));
        }
        [Test] [Explicit]
        public void CalulationValidationExcel()
        {
#if Core
            var dir = AppContext.BaseDirectory;
            dir = Directory.GetParent(dir).Parent.Parent.Parent.FullName;
#else
            var dir = AppDomain.CurrentDomain.BaseDirectory;
#endif
            var pck = new ExcelPackage(new FileInfo(Path.Combine(dir, "Workbooks", "FormulaTest.xlsx")));

            var ws = pck.Workbook.Worksheets["ValidateFormulas"];
            var fr = new Dictionary<string, object>();
            foreach (var cell in ws.Cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    fr.Add(cell.Address, cell.Value);
                }
            }
            pck.Workbook.Calculate();
            var nErrors = 0;
            var errors = new List<Tuple<string, object, object>>();
            foreach (var adr in fr.Keys)
            {
                try
                {
                    if (fr[adr] is double && ws.Cells[adr].Value is double)
                    {
                        var d1 = Convert.ToDouble(fr[adr]);
                        var d2 = Convert.ToDouble(ws.Cells[adr].Value);
                        if (Math.Abs(d1 - d2) < 0.0001)
                        {
                            continue;
                        }
                        else
                        {
                            Assert.That(fr[adr], Is.EqualTo(ws.Cells[adr].Value));
                        }
                    }
                    else
                    {
                        Assert.That(fr[adr], Is.EqualTo(ws.Cells[adr].Value));
                    }
                }
                catch
                {
                    errors.Add(new Tuple<string, object, object>(adr, fr[adr], ws.Cells[adr].Value));
                    nErrors++;
                }
            }
		}
        [Explicit]
        [Test]
        public void TestOneCell()
        {
            var pck = new ExcelPackage(new FileInfo(@"C:\temp\EPPlusTestark\Test4.xlsm"));
            var ws = pck.Workbook.Worksheets.First(); 
            pck.Workbook.Worksheets["Räntebärande formaterat utland"].Cells["M13"].Calculate();
            Assert.That(0d, Is.EqualTo(pck.Workbook.Worksheets["Räntebärande formaterat utland"].Cells["M13"].Value));  
        }
        [Explicit]
        [Test]
        public void TestPrecedence()
        {
            var pck = new ExcelPackage(new FileInfo(@"C:\temp\EPPlusTestark\Precedence.xlsx"));
            var ws = pck.Workbook.Worksheets.Last();
            pck.Workbook.Calculate();
            Assert.That(150d, Is.EqualTo(ws.Cells["A1"].Value));
        }
        [Explicit]
        [Test]
        public void TestDataType()
        {
            var pck = new ExcelPackage(new FileInfo(@"c:\temp\EPPlusTestark\calc_amount.xlsx"));
            var ws = pck.Workbook.Worksheets.First();
            //ws.Names.Add("Name1",ws.Cells["A1"]);
            //ws.Names.Add("Name2", ws.Cells["A2"]);
            ws.Names["PRICE"].Value = 30;
            ws.Names["QUANTITY"].Value = 10;

            ws.Calculate();

            ws.Names["PRICE"].Value = 40;
            ws.Names["QUANTITY"].Value = 20;

            ws.Calculate();
        }
        [Test]
        public void CalcTwiceError()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("CalcTest");
            ws.Names.AddValue("PRICE", 10);
            ws.Names.AddValue("QUANTITY", 11);
            ws.Cells["A1"].Formula="PRICE*QUANTITY";
            ws.Names.AddFormula("AMOUNT", "PRICE*QUANTITY");

            ws.Names["PRICE"].Value = 30;
            ws.Names["QUANTITY"].Value = 10;

            ws.Calculate();
            Assert.That(300D, Is.EqualTo(ws.Cells["A1"].Value));
            Assert.That(300D, Is.EqualTo(ws.Names["AMOUNT"].Value));
            ws.Names["PRICE"].Value = 40;
            ws.Names["QUANTITY"].Value = 20;

            ws.Calculate();
            Assert.That(800D, Is.EqualTo(ws.Cells["A1"].Value));
            Assert.That(800D, Is.EqualTo(ws.Names["AMOUNT"].Value));
        }
        [Test]
        public void IfError()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("CalcTest");
            ws.Cells["A1"].Value = "test1";
            ws.Cells["A5"].Value = "test2";
            ws.Cells["A2"].Value = "Sant";
            ws.Cells["A3"].Value = "Falskt";
            ws.Cells["A4"].Formula = "if(A1>=A5,true,A3)";
            ws.Cells["B1"].Formula = "isText(a1)";
            ws.Cells["B2"].Formula = "isText(\"Test\")";
            ws.Cells["B3"].Formula = "isText(1)";
            ws.Cells["B4"].Formula = "isText(true)";
            ws.Cells["c1"].Formula = "mid(a1,4,15)";

            ws.Calculate();
        }
        [Test]
        public void LeftRightFunctionTest()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("CalcTest");
            ws.SetValue("A1", "asdf");
            ws.Cells["A2"].Formula = "Left(A1, 3)";
            ws.Cells["A3"].Formula = "Left(A1, 10)";
            ws.Cells["A4"].Formula = "Right(A1, 3)";
            ws.Cells["A5"].Formula = "Right(A1, 10)";

            ws.Calculate();
            Assert.That("asd", Is.EqualTo(ws.Cells["A2"].Value));
            Assert.That("asdf", Is.EqualTo(ws.Cells["A3"].Value));
            Assert.That("sdf", Is.EqualTo(ws.Cells["A4"].Value));
            Assert.That("asdf", Is.EqualTo(ws.Cells["A5"].Value));
        }
        [Test]
        public void IfFunctionTest()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("CalcTest");
            ws.SetValue("A1", 123);
            ws.Cells["A2"].Formula = "IF(A1 = 123, 1, -1)";
            ws.Cells["A3"].Formula = "IF(A1 = 1, 1)";
            ws.Cells["A4"].Formula = "IF(A1 = 1, 1, -1)";
            ws.Cells["A5"].Formula = "IF(A1 = 123, 5)";

            ws.Calculate();
            Assert.That(1d, Is.EqualTo(ws.Cells["A2"].Value));
            Assert.That(false, Is.EqualTo(ws.Cells["A3"].Value));
            Assert.That(-1d, Is.EqualTo(ws.Cells["A4"].Value));
            Assert.That(5d, Is.EqualTo(ws.Cells["A5"].Value));
        }
        [Test]
        public void INTFunctionTest()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("CalcTest");
            var currentDate = DateTime.UtcNow.Date;
            ws.SetValue("A1", currentDate.ToString("MM/dd/yyyy"));
            ws.SetValue("A2", currentDate.Date);
            ws.SetValue("A3", "31.1");
            ws.SetValue("A4", 31.1);
            ws.Cells["A5"].Formula = "INT(A1)";
            ws.Cells["A6"].Formula = "INT(A2)";
            ws.Cells["A7"].Formula = "INT(A3)";
            ws.Cells["A8"].Formula = "INT(A4)";

            ws.Calculate();
            Assert.That((int)currentDate.ToOADate(), Is.EqualTo(ws.Cells["A5"].Value));
            Assert.That((int)currentDate.ToOADate(), Is.EqualTo(ws.Cells["A6"].Value));
            Assert.That(31, Is.EqualTo(ws.Cells["A7"].Value));
            Assert.That(31, Is.EqualTo(ws.Cells["A8"].Value));
        }



        public void TestAllWorkbooks()
        {
            StringBuilder sb=new StringBuilder();
            //Add sheets to test in this directory or change it to your testpath.
            string path = @"C:\temp\EPPlusTestark\workbooks";
            if(!Directory.Exists(path)) return;

            foreach (var file in Directory.GetFiles(path, "*.xls*"))
            {
                sb.Append(GetOutput(file));
            }

            if (sb.Length > 0)
            {
                File.WriteAllText(string.Format("TestAllWorkooks{0}.txt", DateTime.Now.ToString("d") + " " + DateTime.Now.ToString("t")), sb.ToString());
                throw(new Exception("Test failed with\r\n\r\n" + sb.ToString()));

            }
        }
		[Test]
		public void CalculateDateMath()
		{
			using (ExcelPackage package = new ExcelPackage())
			{
				var worksheet = package.Workbook.Worksheets.Add("Test");
				var dateCell = worksheet.Cells[2, 2];
				var date = new DateTime(2013, 1, 1);
				dateCell.Value = date;
				var quotedDateCell = worksheet.Cells[2, 3];
				quotedDateCell.Formula = $"\"{date.ToString("d")}\"";
				var dateFormula = "B2";
				var dateFormulaWithMath = "B2+1";
				var quotedDateFormulaWithMath = $"\"{date.ToString("d")}\"+1";
				var quotedDateReferenceFormulaWithMath = "C2+1";
				var expectedDate = 41275.0; // January 1, 2013
				var expectedDateWithMath = 41276.0; // January 2, 2013
				Assert.That(expectedDate, Is.EqualTo(worksheet.Calculate(dateFormula)));
				Assert.That(expectedDateWithMath, Is.EqualTo(worksheet.Calculate(dateFormulaWithMath)));
				Assert.That(expectedDateWithMath, Is.EqualTo(worksheet.Calculate(quotedDateFormulaWithMath)));
				Assert.That(expectedDateWithMath, Is.EqualTo(worksheet.Calculate(quotedDateReferenceFormulaWithMath)));
				var formulaCell = worksheet.Cells[2, 4];
				formulaCell.Formula = dateFormulaWithMath;
				formulaCell.Calculate();
				Assert.That(expectedDateWithMath, Is.EqualTo(formulaCell.Value));
				formulaCell.Formula = quotedDateReferenceFormulaWithMath;
				formulaCell.Calculate();
				Assert.That(expectedDateWithMath, Is.EqualTo(formulaCell.Value));
			}
		}
		private string GetOutput(string file)
        {
            using (var pck = new ExcelPackage(new FileInfo(file)))
            {
                var fr = new Dictionary<string, object>();
                foreach (var ws in pck.Workbook.Worksheets)
                {
                    if (!(ws is ExcelChartsheet))
                    {
                        foreach (var cell in ws.Cells)
                        {
                            if (!string.IsNullOrEmpty(cell.Formula))
                            {
                                fr.Add(ws.PositionID.ToString() + "," + cell.Address, cell.Value);
                                ws.SetValueInner(cell.Start.Row, cell.Start.Column, null);
                            }
                        }
                    }
                }

                pck.Workbook.Calculate();
                var nErrors = 0;
                var errors = new List<Tuple<string, object, object>>();
                ExcelWorksheet sheet=null;
                string adr="";
                var fileErr = new System.IO.StreamWriter(new FileStream("c:\\temp\\err.txt",FileMode.Append));
                foreach (var cell in fr.Keys)
                {
                    try
                    {
                        var spl = cell.Split(',');
                        var ix = int.Parse(spl[0]);
                        sheet = pck.Workbook.Worksheets[ix];
                        adr = spl[1];
                        if (fr[cell] is double && (sheet.Cells[adr].Value is double || sheet.Cells[adr].Value is decimal  || OfficeOpenXml.Compatibility.TypeCompat.IsPrimitive(sheet.Cells[adr].Value)))
                        {
                            var d1 = Convert.ToDouble(fr[cell]);
                            var d2 = Convert.ToDouble(sheet.Cells[adr].Value);
                            //if (Math.Abs(d1 - d2) < double.Epsilon)
                            if(double.Equals(d1,d2))
                            {
                                continue;
                            }
                            else
                            {
                                //errors.Add(new Tuple<string, object, object>(adr, fr[cell], sheet.Cells[adr].Value));
                                fileErr.WriteLine("Diff cell " + sheet.Name + "!" + adr +"\t" + d1.ToString("R15") + "\t" + d2.ToString("R15"));
                            }
                        }
                        else
                        {
                            if ((fr[cell]??"").ToString() != (sheet.Cells[adr].Value??"").ToString())
                            {
                                fileErr.WriteLine("String?  cell " + sheet.Name + "!" + adr + "\t" + (fr[cell] ?? "").ToString() + "\t" + (sheet.Cells[adr].Value??"").ToString());
                            }
                            //errors.Add(new Tuple<string, object, object>(adr, fr[cell], sheet.Cells[adr].Value));
                        }
                    }
                    catch (Exception e)
                    {                        
                        fileErr.WriteLine("Exception cell " + sheet.Name + "!" + adr + "\t" + fr[cell].ToString() + "\t" + sheet.Cells[adr].Value +  "\t" + e.Message);
                        fileErr.WriteLine("***************************");
                        fileErr.WriteLine(e.ToString());
                        nErrors++;
                    }
                }
                fileErr.Close();
                return nErrors.ToString();
            }
        }
    }
}
