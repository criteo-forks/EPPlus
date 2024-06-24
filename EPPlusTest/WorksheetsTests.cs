using System;
using System.IO;
using NUnit.Framework;
using OfficeOpenXml;
using System.Reflection;

namespace EPPlusTest
{
	[TestFixture]
	public class WorksheetsTests
	{
		private ExcelPackage package;
		private ExcelWorkbook workbook;

		[SetUp]
		public void TestInitialize()
		{
			package = new ExcelPackage();
			workbook = package.Workbook;
			workbook.Worksheets.Add("NEW1");
		}

		[Test]
		public void ConfirmFileStructure()
		{
			Assert.That(package, Is.Not.Null, "Package not created");
			Assert.That(workbook, Is.Not.Null, "No workbook found");
		}

		[Test]
		public void ShouldBeAbleToDeleteAndThenAdd()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Delete(1);
			workbook.Worksheets.Add("NEW3");
		}

		[Test]
		public void DeleteByNameWhereWorkSheetExists()
		{
		    workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Delete("NEW2");
        }

		[Test]
		public void DeleteByNameWhereWorkSheetDoesNotExist()
		{
			Assert.Throws<ArgumentException>(() =>
			{
				workbook.Worksheets.Add("NEW2");
				workbook.Worksheets.Delete("NEW3");
			});
		}

		[Test]
		public void MoveBeforeByNameWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveBefore("NEW4", "NEW2");

			CompareOrderOfWorksheetsAfterSaving(package);
		}

		[Test]
		public void MoveAfterByNameWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveAfter("NEW4", "NEW2");

			CompareOrderOfWorksheetsAfterSaving(package);
		}

		[Test]
		public void MoveBeforeByPositionWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveBefore(4, 2);

			CompareOrderOfWorksheetsAfterSaving(package);
		}

		[Test]
		public void MoveAfterByPositionWhereWorkSheetExists()
		{
			workbook.Worksheets.Add("NEW2");
			workbook.Worksheets.Add("NEW3");
			workbook.Worksheets.Add("NEW4");
			workbook.Worksheets.Add("NEW5");

			workbook.Worksheets.MoveAfter(4, 2);

			CompareOrderOfWorksheetsAfterSaving(package);
		}
        #region Delete Column with Save Tests

        private const string OutputDirectory = @"d:\temp\";

        [Explicit]
        [Test]
        public void DeleteFirstColumnInRangeColumnShouldBeDeleted()
        {
            // Arrange
            ExcelPackage pck = new ExcelPackage();
            using (
                Stream file =
              Assembly.GetExecutingAssembly().GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
            {
                pck.Load(file);
            }
            var wsData = pck.Workbook.Worksheets[1];

            // Act
            wsData.DeleteColumn(1);
            pck.SaveAs(new FileInfo(OutputDirectory + "AfterDeleteColumn.xlsx"));

            // Assert
            Assert.That("Title", Is.EqualTo(wsData.Cells["A1"].Text));
            Assert.That("First Name", Is.EqualTo(wsData.Cells["B1"].Text));
            Assert.That("Family Name", Is.EqualTo(wsData.Cells["C1"].Text));
        }


        [Test] [Explicit]
        public void DeleteLastColumnInRangeColumnShouldBeDeleted()
        {
            // Arrange
            ExcelPackage pck = new ExcelPackage();
            using (
                Stream file =
              Assembly.GetExecutingAssembly().GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
            {
                pck.Load(file);
            }
            var wsData = pck.Workbook.Worksheets[1];

            // Act
            wsData.DeleteColumn(4);
            pck.SaveAs(new FileInfo(OutputDirectory + "AfterDeleteColumn.xlsx"));

            // Assert
            Assert.That("Id", Is.EqualTo(wsData.Cells["A1"].Text));
            Assert.That("Title", Is.EqualTo(wsData.Cells["B1"].Text));
            Assert.That("First Name", Is.EqualTo(wsData.Cells["C1"].Text));
        }

        [Test] [Explicit]
        public void DeleteColumnAfterNormalRangeSheetShouldRemainUnchanged()
        {
            // Arrange
            ExcelPackage pck = new ExcelPackage();
            using (
                Stream file =
              Assembly.GetExecutingAssembly().GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
            {
                pck.Load(file);
            }
            var wsData = pck.Workbook.Worksheets[1];

            // Act
            wsData.DeleteColumn(5);
            pck.SaveAs(new FileInfo(OutputDirectory + "AfterDeleteColumn.xlsx"));

            // Assert
            Assert.That("Id", Is.EqualTo(wsData.Cells["A1"].Text));
            Assert.That("Title", Is.EqualTo(wsData.Cells["B1"].Text));
            Assert.That("First Name", Is.EqualTo(wsData.Cells["C1"].Text));
            Assert.That("Family Name", Is.EqualTo(wsData.Cells["D1"].Text));

        }

        [Test] [Explicit]
        public void DeleteColumnBeforeRangeMimitThrowsArgumentException()
        {
	        Assert.Throws<ArgumentException>(() =>
	        {
		        // Arrange
		        ExcelPackage pck = new ExcelPackage();
		        using (
			        Stream file =
			        Assembly.GetExecutingAssembly()
				        .GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
		        {
			        pck.Load(file);
		        }

		        var wsData = pck.Workbook.Worksheets[1];

		        // Act
		        wsData.DeleteColumn(0);

		        // Assert
		        Assert.Fail();
	        });
        }

        [Test] [Explicit]
        public void DeleteColumnAfterRangeLimitThrowsArgumentException()
        {
	        Assert.Throws<ArgumentException>(() =>
		    {
	            // Arrange
	            ExcelPackage pck = new ExcelPackage();
	            using (
	                Stream file =
	              Assembly.GetExecutingAssembly().GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
	            {
	                pck.Load(file);
	            }
	            var wsData = pck.Workbook.Worksheets[1];

	            // Act
	            wsData.DeleteColumn(16385);

	            // Assert
	            Assert.Fail();
		    });
        }

        [Test] [Explicit]
        public void DeleteFirstTwoColumnsFromRangeColumnsShouldBeDeleted()
        {
            // Arrange
            ExcelPackage pck = new ExcelPackage();
            using (
                Stream file =
              Assembly.GetExecutingAssembly().GetManifestResourceStream("EPPlusTest.TestWorkbooks.PreDeleteColumn.xls"))
            {
                pck.Load(file);
            }
            var wsData = pck.Workbook.Worksheets[1];

            // Act
            wsData.DeleteColumn(1, 2);
            pck.SaveAs(new FileInfo(OutputDirectory + "AfterDeleteColumn.xlsx"));

            // Assert
            Assert.That("First Name", Is.EqualTo(wsData.Cells["A1"].Text));
            Assert.That("Family Name", Is.EqualTo(wsData.Cells["B1"].Text));

        }
        #endregion

        [Test]
        public void RangeClearMethodShouldNotClearSurroundingCells()
        {
            var wks  = workbook.Worksheets.Add("test");
            wks.Cells[2, 2].Value = "something";
            wks.Cells[2, 3].Value = "something";

            wks.Cells[2, 3].Clear();

            Assert.That(wks.Cells[2, 2].Value, Is.Not.Null);
            Assert.Equals("something", wks.Cells[2, 2].Value);
            Assert.Equals(null, wks.Cells[2, 3].Value);
        }

        private static void CompareOrderOfWorksheetsAfterSaving(ExcelPackage editedPackage)
		{
			var packageStream = new MemoryStream();
			editedPackage.SaveAs(packageStream);

			var newPackage = new ExcelPackage(packageStream);
            newPackage.Compatibility.IsWorksheets1Based = true;
            var positionId = 1;
			foreach (var worksheet in editedPackage.Workbook.Worksheets)
			{
				Assert.That(worksheet.Name, Is.EqualTo(newPackage.Workbook.Worksheets[positionId].Name));
				positionId++;
			}
		}
	}
}
