using System;
using NUnit.Framework;
using OfficeOpenXml;

namespace EPPlusTest
{
    [TestFixture]
    public class ExcelCellBaseTest
    {
        #region UpdateFormulaReferences Tests
        [Test]
        public void UpdateFormulaReferencesOnTheSameSheet()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("C3", 3, 3, 2, 2, "sheet", "sheet");
            Assert.That("F6", Is.EqualTo(result));
        }

        [Test]
        public void UpdateFormulaReferencesIgnoresIncorrectSheet()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("C3", 3, 3, 2, 2, "sheet", "other sheet");
            Assert.That("C3", Is.EqualTo(result));
        }

        [Test]
        public void UpdateFormulaReferencesFullyQualifiedReferenceOnTheSameSheet()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("'sheet name here'!C3", 3, 3, 2, 2, "sheet name here", "sheet name here");
            Assert.That("'sheet name here'!F6", Is.EqualTo(result));
        }

        [Test]
        public void UpdateFormulaReferencesFullyQualifiedCrossSheetReferenceArray()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("SUM('sheet name here'!B2:D4)", 3, 3, 3, 3, "cross sheet", "sheet name here");
            Assert.That("SUM('sheet name here'!B2:G7)", Is.EqualTo(result));
        }

        [Test]
        public void UpdateFormulaReferencesFullyQualifiedReferenceOnADifferentSheet()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("'updated sheet'!C3", 3, 3, 2, 2, "boring sheet", "updated sheet");
            Assert.That("'updated sheet'!F6", Is.EqualTo(result));
        }

        [Test]
        public void UpdateFormulaReferencesReferencingADifferentSheetIsNotUpdated()
        {
            var result = ExcelCellBase.UpdateFormulaReferences("'boring sheet'!C3", 3, 3, 2, 2, "boring sheet", "updated sheet");
            Assert.That("'boring sheet'!C3", Is.EqualTo(result));
        }
        #endregion

        #region UpdateCrossSheetReferenceNames Tests
        [Test]
        public void UpdateFormulaSheetReferences()
        {
          var result = ExcelCellBase.UpdateFormulaSheetReferences("5+'OldSheet'!$G3+'Some Other Sheet'!C3+SUM(1,2,3)", "OldSheet", "NewSheet");
          Assert.Equals("5+'NewSheet'!$G3+'Some Other Sheet'!C3+SUM(1,2,3)", result);
        }

        [Test]
        public void UpdateFormulaSheetReferencesNullOldSheetThrowsException()
        {
           Assert.Throws<ArgumentNullException>(() => ExcelCellBase.UpdateFormulaSheetReferences("formula", null, "sheet2"));
        }

        [Test]
        public void UpdateFormulaSheetReferencesEmptyOldSheetThrowsException()
        {
            Assert.Throws<ArgumentNullException>(() => ExcelCellBase.UpdateFormulaSheetReferences("formula", string.Empty, "sheet2"));
        }

        [Test]
        public void UpdateFormulaSheetReferencesNullNewSheetThrowsException()
        {
            Assert.Throws<ArgumentNullException>(() => ExcelCellBase.UpdateFormulaSheetReferences("formula", "sheet1", null));
        }

        [Test]
        public void UpdateFormulaSheetReferencesEmptyNewSheetThrowsException()
        {
            Assert.Throws<ArgumentNullException>(() => ExcelCellBase.UpdateFormulaSheetReferences("formula", "sheet1", string.Empty));
        }
        #endregion
  }
}
