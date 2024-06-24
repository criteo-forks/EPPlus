using System;
using System.Text;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.ExpressionGraph;
using FakeItEasy;
using System.IO;
using System.Threading;

namespace EPPlusTest.FormulaParsing.IntegrationTests.BuiltInFunctions
{
    [TestFixture]
    public class DateAndTimeFunctionsTests : FormulaParserTestBase
    {
        [SetUp]
        public void Setup()
        {
            var excelDataProvider = A.Fake<ExcelDataProvider>();
            _parser = new FormulaParser(excelDataProvider);
        }

        [Test]
        public void DateShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Date(2012, 2, 2)");
            Assert.That(new DateTime(2012, 2, 2).ToOADate(), Is.EqualTo(result));
        }

        [Test]
        public void DateShouldHandleCellReference()
        {
            using (var pck = new ExcelPackage())
            {
                var sheet = pck.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = 2012d;
                sheet.Cells["A2"].Formula = "Date(A1, 2, 2)";
                sheet.Calculate();
                var result = sheet.Cells["A2"].Value;
                Assert.That(new DateTime(2012, 2, 2).ToOADate(), Is.EqualTo(result));
            }

        }

        [Test]
        public void TodayShouldReturnAResult()
        {
            var result = _parser.Parse("Today()");
            Assert.That(DateTime.FromOADate((double)result), Is.InstanceOf<DateTime>());
        }

        [Test]
        public void NowShouldReturnAResult()
        {
            var result = _parser.Parse("now()");
            Assert.That(DateTime.FromOADate((double)result), Is.InstanceOf<DateTime>());
        }

        [Test]
        public void DayShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Day(Date(2012, 4, 2))");
            Assert.That(2, Is.EqualTo(result));
        }

        [Test]
        public void MonthShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Month(Date(2012, 4, 2))");
            Assert.That(4, Is.EqualTo(result));
        }

        [Test]
        public void YearShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Year(Date(2012, 2, 2))");
            Assert.That(2012, Is.EqualTo(result));
        }

        [Test]
        public void TimeShouldReturnCorrectResult()
        {
            var expectedResult = ((double)(12 * 60 * 60 + 13 * 60 + 14))/((double)(24 * 60 * 60));
            var result = _parser.Parse("Time(12, 13, 14)");
            Assert.That(expectedResult, Is.EqualTo(result));
        }

        [Test]
        public void HourShouldReturnCorrectResult()
        {
            var result = _parser.Parse("HOUR(Time(12, 13, 14))");
            Assert.That(12, Is.EqualTo(result));
        }

        [Test]
        public void MinuteShouldReturnCorrectResult()
        {
            var result = _parser.Parse("minute(Time(12, 13, 14))");
            Assert.That(13, Is.EqualTo(result));
        }

        [Test]
        public void SecondShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Second(Time(12, 13, 59))");
            Assert.That(59, Is.EqualTo(result));
        }

        [Test]
        public void SecondShouldReturnCorrectResultWhenParsingString()
        {
            var result = _parser.Parse("Second(\"10:12:14\")");
            Assert.That(14, Is.EqualTo(result));
        }

        [Test]
        public void MinuteShouldReturnCorrectResultWhenParsingString()
        {
            var result = _parser.Parse("Minute(\"10:12:14 AM\")");
            Assert.That(12, Is.EqualTo(result));
        }

        [Test]
        public void HourShouldReturnCorrectResultWhenParsingString()
        {
            var result = _parser.Parse("Hour(\"10:12:14\")");
            Assert.That(10, Is.EqualTo(result));
        }

        [Test]
        public void Day360ShouldReturnCorrectResult()
        {
            var result = _parser.Parse("Days360(Date(2012, 4, 2), Date(2012, 5, 2))");
            Assert.That(30, Is.EqualTo(result));
        }

        [Test]
        public void YearfracShouldReturnAResult()
        {
            var result = _parser.Parse("Yearfrac(Date(2012, 4, 2), Date(2012, 5, 2))");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void IsoWeekNumShouldReturnAResult()
        {
            var result = _parser.Parse("IsoWeekNum(Date(2012, 4, 2))");
            Assert.That(result, Is.InstanceOf<int>());
        }

        [Test]
        public void EomonthShouldReturnAResult()
        {
            var result = _parser.Parse("Eomonth(Date(2013, 2, 2), 3)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void WorkdayShouldReturnAResult()
        {
            var result = _parser.Parse("Workday(Date(2013, 2, 2), 3)");
            Assert.That(result, Is.InstanceOf<double>());
        }

        [Test]
        public void DateNotEqualToStringShouldBeTrue()
        {
            var result = _parser.Parse("TODAY() <> \"\"");
            Assert.That((bool)result);
        }

        [Test]
        public void Calculation5()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = "John";
            ws.Cells["B1"].Value = "Doe";
            ws.Cells["C1"].Formula = "B1&\", \"&A1";
            ws.Calculate();
            Assert.Equals("Doe, John", ws.Cells["C1"].Value);
        }

        [Test]
        public void HourWithExcelReference()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = new DateTime(2014, 1, 1, 10, 11, 12).ToOADate();
            ws.Cells["B1"].Formula = "HOUR(A1)";
            ws.Calculate();
            Assert.That(10, Is.EqualTo(ws.Cells["B1"].Value));
        }

        [Test]
        public void MinuteWithExcelReference()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = new DateTime(2014, 1, 1, 10, 11, 12).ToOADate();
            ws.Cells["B1"].Formula = "MINUTE(A1)";
            ws.Calculate();
            Assert.That(11, Is.EqualTo(ws.Cells["B1"].Value));
        }

        [Test]
        public void SecondWithExcelReference()
        {
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = new DateTime(2014, 1, 1, 10, 11, 12).ToOADate();
            ws.Cells["B1"].Formula = "SECOND(A1)";
            ws.Calculate();
            Assert.That(12, Is.EqualTo(ws.Cells["B1"].Value));
        }
#if (!Core)
        [Test]
        public void DateValueTest1()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            ws.Cells["A1"].Value = "21 JAN 2015";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.That(new DateTime(2015, 1, 21).ToOADate(), Is.EqualTo(ws.Cells["B1"].Value));
        }

        [Test]
        public void DateValueTestWithoutYear()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var currentYear = DateTime.Now.Year;
            ws.Cells["A1"].Value = "21 JAN";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.That(new DateTime(currentYear, 1, 21).ToOADate(), Is.EqualTo(ws.Cells["B1"].Value));
        }

        [Test]
        public void DateValueTestWithTwoDigitYear()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var expectedYear = 1930;
            ws.Cells["A1"].Value = "01/01/30";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.That(new DateTime(expectedYear, 1, 1).ToOADate(), Is.EqualTo(ws.Cells["B1"].Value));
        }

        [Test]
        public void DateValueTestWithTwoDigitYear2()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var expectedYear = 2029;
            ws.Cells["A1"].Value = "01/01/29";
            ws.Cells["B1"].Formula = "DateValue(A1)";
            ws.Calculate();
            Assert.That(new DateTime(expectedYear, 1, 1).ToOADate(), Is.EqualTo(ws.Cells["B1"].Value));
        }


        [Test]
        public void TimeValueTestPm()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var currentYear = DateTime.Now.Year;
            ws.Cells["A1"].Value = "2:23 pm";
            ws.Cells["B1"].Formula = "TimeValue(A1)";
            ws.Calculate();
            var result = (double) ws.Cells["B1"].Value;
            Assert.Equals(0.599, Math.Round(result, 3));
        }


        [Test]
        public void TimeValueTestFullDate()
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            var pck = new ExcelPackage();
            var ws = pck.Workbook.Worksheets.Add("Calc1");
            var currentYear = DateTime.Now.Year;
            ws.Cells["A1"].Value = "01/01/2011 02:23";
            ws.Cells["B1"].Formula = "TimeValue(A1)";
            ws.Calculate();
            var result = (double)ws.Cells["B1"].Value;
            Assert.Equals(0.099, Math.Round(result, 3));
        }
#endif
    }
}
