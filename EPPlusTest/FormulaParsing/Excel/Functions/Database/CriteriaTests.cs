using System;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Database
{
    [TestFixture]
    public class CriteriaTests
    {
        [Test]
        public void CriteriaShouldReadFieldsAndValues()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Crit1";
                sheet.Cells["B1"].Value = "Crit2";
                sheet.Cells["A2"].Value = 1;
                sheet.Cells["B2"].Value = 2;

                var provider = new EpplusExcelDataProvider(package);

                var criteria = new ExcelDatabaseCriteria(provider, "A1:B2");

                Assert.That(2, Is.EqualTo(criteria.Items.Count));
                Assert.That("crit1", Is.EqualTo(criteria.Items.Keys.First().ToString()));
                Assert.That("crit2", Is.EqualTo(criteria.Items.Keys.Last().ToString()));
                Assert.That(1, Is.EqualTo(criteria.Items.Values.First()));
                Assert.That(2, Is.EqualTo(criteria.Items.Values.Last()));
            }
        }

        [Test]
        public void CriteriaShouldIgnoreEmptyFields1()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Crit1";
                sheet.Cells["B1"].Value = "Crit2";
                sheet.Cells["A2"].Value = 1;

                var provider = new EpplusExcelDataProvider(package);

                var criteria = new ExcelDatabaseCriteria(provider, "A1:B2");

                Assert.That(1, Is.EqualTo(criteria.Items.Count));
                Assert.That("crit1", Is.EqualTo(criteria.Items.Keys.First().ToString()));
                Assert.That(1, Is.EqualTo(criteria.Items.Values.Last()));
            }
        }

        [Test]
        public void CriteriaShouldIgnoreEmptyFields2()
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("test");
                sheet.Cells["A1"].Value = "Crit1";
                sheet.Cells["A2"].Value = 1;

                var provider = new EpplusExcelDataProvider(package);

                var criteria = new ExcelDatabaseCriteria(provider, "A1:B2");

                Assert.That(1, Is.EqualTo(criteria.Items.Count));
                Assert.That("crit1", Is.EqualTo(criteria.Items.Keys.First().ToString()));
                Assert.That(1, Is.EqualTo(criteria.Items.Values.Last()));
            }
        }

    }
}
