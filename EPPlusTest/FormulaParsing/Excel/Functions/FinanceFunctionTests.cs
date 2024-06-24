using System;
using NUnit.Framework;
using OfficeOpenXml;

namespace EPPlusTest.FormulaParsing.Excel.Functions
{
    [TestFixture]
    public class FinanceFunctionTests
    {
        [Test]
        public void PmtTest1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Formula = "PMT( 5%/12, 60, 50000 )";
                sheet.Calculate();
                var value = sheet.Cells["A1"].Value;
                var value2 = System.Math.Round(Convert.ToDouble(value), 2);
                Assert.That(-943.56, Is.EqualTo(value2));
            }
        }
    }
}
