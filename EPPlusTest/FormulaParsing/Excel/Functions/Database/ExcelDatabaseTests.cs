using System;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Database;

namespace EPPlusTest.FormulaParsing.Excel.Functions.Database
{
    [TestFixture]
    public class ExcelDatabaseTests
    {
        [Test]
        public void DatabaseShouldReadFields()
        {
            using (var package = new ExcelPackage())
            {
                var database = GetDatabase(package);

                Assert.That(2, Is.EqualTo(database.Fields.Count()), "count was not 2");
                Assert.That("col1", Is.EqualTo(database.Fields.First().FieldName), "first fieldname was not 'col1'");
                Assert.That("col2", Is.EqualTo(database.Fields.Last().FieldName), "last fieldname was not 'col12'");
            }
        }

        [Test]
        public void HasMoreRowsShouldBeTrueWhenInitialized()
        {
            using (var package = new ExcelPackage())
            {
                var database = GetDatabase(package);

                Assert.That(database.HasMoreRows);
            }
            
        }

        [Test]
        public void HasMoreRowsShouldBeFalseWhenLastRowIsRead()
        {
            using (var package = new ExcelPackage())
            {
                var database = GetDatabase(package);
                database.Read();

                Assert.That(!database.HasMoreRows);
            }

        }

        [Test]
        public void DatabaseShouldReadFieldsInRow()
        {
            using (var package = new ExcelPackage())
            {
                var database = GetDatabase(package);
                var row = database.Read();

                Assert.That(1, Is.EqualTo(row["col1"]));
                Assert.That(2, Is.EqualTo(row["col2"]));
            }

        }

        private static ExcelDatabase GetDatabase(ExcelPackage package)
        {
            var provider = new EpplusExcelDataProvider(package);
            var sheet = package.Workbook.Worksheets.Add("test");
            sheet.Cells["A1"].Value = "col1";
            sheet.Cells["A2"].Value = 1;
            sheet.Cells["B1"].Value = "col2";
            sheet.Cells["B2"].Value = 2;
            var database = new ExcelDatabase(provider, "A1:B2");
            return database;
        }
    }
}
